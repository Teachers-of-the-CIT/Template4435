using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

using Excel = Microsoft.Office.Interop.Excel;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private clientEntities _clientEntities;
        public MainWindow()
        {
            InitializeComponent();
            _clientEntities = new clientEntities();
        }


        private void BtnCHELNY_Click(object sender, RoutedEventArgs e)
        {



        }
        private void BtnNikita_Click(object sender, RoutedEventArgs e)
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
            using (clientEntities lR2Entities = new clientEntities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    lR2Entities.client.Add(new client()
                    {
                        id = i,
                        Kod_zakaz = list[i, 1],
                        Date = list[i, 2],
                        Time = list[i, 3],
                        Kod_client = list[i, 4],
                        Uslugi = list[i, 5],
                        Status = list[i, 6],
                        Date_zakrit = list[i, 7],
                        Vremya = i
                    });
                }
                lR2Entities.SaveChanges();
            }
        }

        private void Maximov_Click(object sender, RoutedEventArgs e)
        {
            List<client> allzakaz;


            var alltime = _clientEntities.client.Select(u => new { Vremya = u.Vremya }).Distinct().ToList();
            allzakaz = _clientEntities.client.ToList().OrderBy(g => g.Kod_zakaz).ToList();

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = alltime.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);



            for (int i = 0; i < alltime.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
     
                worksheet.Name = Convert.ToString(alltime[i].Vremya);


                worksheet.Cells[1][2][startRowIndex] = "id";
                worksheet.Cells[2][2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][2][startRowIndex] = "Дата создания";
                worksheet.Cells[4][2][startRowIndex] = "Код клиента";
                worksheet.Cells[5][2][startRowIndex] = "Услуги";
                startRowIndex++;

                var studentsCategories = allzakaz.GroupBy(s => s.Vremya).ToList();
                foreach (var students in studentsCategories)
                {
                    if (Convert.ToInt32(students.Key) == alltime[i].Vremya)
                    {
                        Excel.Range headerRange =
                        worksheet.Range[worksheet.Cells[1][1],
                        worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = alltime[i].Vremya;
                        headerRange.HorizontalAlignment =
                        Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;
                        foreach (client student in allzakaz)
                        {
                            if (Convert.ToInt32(student.Vremya) == students.Key)
                            {
                                worksheet.Cells[1][startRowIndex] =
                                student.id;
                                worksheet.Cells[2][startRowIndex] =
                                student.Kod_zakaz;
                                worksheet.Cells[3][startRowIndex] =
                                student.Date;
                                worksheet.Cells[4][startRowIndex] =
                                student.Kod_client;
                                worksheet.Cells[5][startRowIndex] =
                                student.Uslugi;

                                startRowIndex++;
                            }

                        }
                    }

                    else
                    {
                        continue;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 2]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = 
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = 
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;





        }







        private void Adieva_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Адиева Айгуль Ринатовна", "4435_Адиева_Айгуль");
        }
    }
}

