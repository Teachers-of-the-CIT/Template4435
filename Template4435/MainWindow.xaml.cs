using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Entities _entities;
        public MainWindow()
        {
            InitializeComponent();
            _entities = new Entities();
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (Entities entities = new Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    entities.users.Add(new users()
                    {
                        Id = list[i, 0],
                        FIO = list[i, 1],
                        Login = list[i, 2],
                        Doljnost = list[i, 3]
                    });
                }
                entities.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<users> allUsers;
            
            var allDoljnosti = _entities.users.Select(u => new {Doljnost = u.Doljnost}).Distinct().ToList();
            allUsers = _entities.users.ToList().OrderBy(g => g.Id).ToList();
            
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allDoljnosti.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allDoljnosti.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(allDoljnosti[i].Doljnost);
               
                    worksheet.Cells[1][2] = "Порядковый номер";
                    worksheet.Cells[2][2] = "ФИО";
                    worksheet.Cells[3][2] = "Логин";
                startRowIndex++;    
                var usersCategories = allUsers.GroupBy(s => s.Doljnost).ToList();
                foreach (var users in usersCategories)
                {
                    if (users.Key == allDoljnosti[i].Doljnost)
                    {
                        Excel.Range headerRange =
                        worksheet.Range[worksheet.Cells[1][1],
                        worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = allDoljnosti[i].Doljnost;
                        headerRange.HorizontalAlignment =
                        Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;
                        foreach (users user in allUsers)
                        {
                            if (user.Doljnost == users.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = user.Id;
                                worksheet.Cells[2][startRowIndex] = user.FIO;
                                worksheet.Cells[3][startRowIndex] = user.Login;
                                startRowIndex++;
                            }
                        }
                        worksheet.Cells[1][startRowIndex].Formula =
                        $"=СЧЁТ(A3:A{startRowIndex - 1})";
                        worksheet.Cells[1][startRowIndex].Font.Bold =
                        true;
                    }
                    else
                    {
                        continue;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1],
                    worksheet.Cells[3][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        
        }
    }
}
