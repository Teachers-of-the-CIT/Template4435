using System.Windows;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using System.IO;
using Template4435.KhaertdinovDG;
using System.Linq;
using Microsoft.Win32;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для KhaertdinovDGWindow.xaml
    /// </summary>
    public partial class KhaertdinovDGWindow : System.Windows.Window
    {
        //(localdb)\MSSqlLocalDB
        public KhaertdinovDGWindow()
        {
            InitializeComponent();

        }

        string root = @"C:\Users\LenyaPlay\Downloads\";
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = root;
            if (!ofd.ShowDialog().GetValueOrDefault())
            {
                return;
            }

            var filename = ofd.FileName;

            var app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = app.Workbooks.Open(filename);
            wkb.SaveAs(root + "1.txt", XlFileFormat.xlUnicodeText);
            wkb.Close();
            string[] lines = File.ReadAllLines(root + "1.txt");
            File.Delete(root + "1.txt");
            

            using(KhaertdinovModelContainer db = new KhaertdinovModelContainer())
            {
                foreach (var line in lines)
                {
                    if (line == lines[0])
                        continue;
                    var cells = line.Split('\t');
                    string TypeName = cells[2];
                    Type type = new Type() { Name = TypeName };
                    if (db.TypeSet.Where(x => x.Name == TypeName).Count() == 0)
                    {
                        db.TypeSet.Add(type);
                        db.SaveChanges();
                    }
                    else
                        type = db.TypeSet.Where(x => x.Name == TypeName).First();

                    Service service = new Service()
                    {
                        Name = cells[1],
                        Type = type,
                        Code = cells[3],
                        Price = int.Parse(cells[4])
                    };
                    db.ServiceSet.Add(service);
                }
                db.SaveChanges();

            }

        }

        private void BtmExport_Click(object sender, RoutedEventArgs e)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wkb = app.Workbooks.Add();

            using (KhaertdinovModelContainer db = new KhaertdinovModelContainer())
            {
                var sets = db.ServiceSet.GroupBy(s => s.Type);
                
                foreach (var g in sets)
                {
                    Worksheet worksheet = wkb.Sheets.Add();
                    worksheet.Name = g.First().Type.Name;

                    var row = 1;
                    worksheet.Rows[row].EntireRow.Font.Bold = true;
                    worksheet.Cells[row, "A"] = "Id";
                    worksheet.Cells[row, "B"] = "Название услуги";
                    worksheet.Cells[row, "C"] = "Стоимость";
                    
                    foreach (var s in g.OrderBy(x => x.Price))
                    {
                        row++;
                        worksheet.Cells[row, "A"] = s.Id;
                        worksheet.Cells[row, "B"] = s.Name;
                        worksheet.Cells[row, "C"] = s.Price;
                    }
                }
            }

            wkb.Sheets[wkb.Sheets.Count].Delete();

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = root;
            sfd.Filter = "Excel (*.xlsx)|.xlsx";
            if (!sfd.ShowDialog().GetValueOrDefault())
            {
                return;
            }
            wkb.SaveAs(sfd.FileName);
            wkb.Close();

        }
    }
}
