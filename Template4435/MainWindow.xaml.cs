using Microsoft.Win32;
using System.Text.Json;
using System.Text.Json.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Window = System.Windows.Window;
using System.IO;


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
                        Vremya = list[i, 8]
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
                    if (students.Key == alltime[i].Vremya)
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
                            if (student.Vremya == students.Key)
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

        private void Word_Click(object sender, RoutedEventArgs e)
        {

            List<client> allzakaz;


            var alltime = _clientEntities.client.Select(u => new { Vremya = u.Vremya }).Distinct().ToList();
            allzakaz = _clientEntities.client.ToList().OrderBy(g => g.Kod_zakaz).ToList();
            var studentsCategories = allzakaz.GroupBy(s => s.Vremya).ToList();
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();
            foreach (var group in studentsCategories)
            {
                Word.Paragraph paragraph =
                document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = Convert.ToString(alltime.Where(g => g.Vremya == group.Key).FirstOrDefault().Vremya);
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table studentsTable =
                document.Tables.Add(tableRange, group.Count() + 1, 5);
                studentsTable.Borders.InsideLineStyle = studentsTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                studentsTable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;
                cellRange = studentsTable.Cell(1, 1).Range;
                cellRange.Text = "id";
                cellRange = studentsTable.Cell(1, 2).Range;
                cellRange.Text = "Код заказа";
                cellRange = studentsTable.Cell(1, 3).Range;
                cellRange.Text = "Дата";
                cellRange = studentsTable.Cell(1, 4).Range;
                cellRange.Text = "Код клиента";
                cellRange = studentsTable.Cell(1, 5).Range;
                cellRange.Text = "Услуги";
                studentsTable.Rows[1].Range.Bold = 1;
                studentsTable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                int i = 1;
                foreach (var currentStudent in group)
                {
                    cellRange = studentsTable.Cell(i + 1, 1).Range;
                    cellRange.Text = currentStudent.id.ToString();
                    cellRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = studentsTable.Cell(i + 1, 2).Range;
                    cellRange.Text = currentStudent.Kod_zakaz;
                    cellRange.ParagraphFormat.Alignment =
                    WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = studentsTable.Cell(i + 1, 3).Range;
                    cellRange.Text = currentStudent.Date;
                    cellRange.ParagraphFormat.Alignment =
                    WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = studentsTable.Cell(i + 1, 4).Range;
                    cellRange.Text = currentStudent.Kod_client;
                    cellRange.ParagraphFormat.Alignment =
                    WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = studentsTable.Cell(i + 1, 5).Range;
                    cellRange.Text = currentStudent.Uslugi;
                    cellRange.ParagraphFormat.Alignment =
                    WdParagraphAlignment.wdAlignParagraphCenter;
                    i++;
                    document.Words.Last.InsertBreak(WdBreakType.wdPageBreak);
                }


            }
            app.Visible = true;

            document.SaveAs2(@"C:\Users\Aigul\Documents\outputFileWord.docx");
            document.SaveAs2(@"C:\Users\Aigul\Documents\outputFilePdf.pdf",
            WdExportFormat.wdExportFormatPDF);

        }

        private async void json_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Title = "Выберите файл JSON"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            using (FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate))
            {
                List<client> users = await JsonSerializer.DeserializeAsync<List<client>>(fs);
                
            
           

                using (clientEntities lR2Entities = new clientEntities())
                {
                    foreach (client user in users)
                    {
                        client userr = new client();
                        userr.id = user.id ;
                        userr.Kod_zakaz = user.Kod_zakaz;
                        userr.Date = user.Date;
                        userr.Time = user.Time;
                        userr.Kod_client = user.Kod_client;
                        userr.Uslugi = user.Uslugi;
                        userr.Status = user.Status;
                        userr.Date_zakrit = user.Date_zakrit;
                        userr.Vremya = user.Vremya;

                        lR2Entities.client.Add(userr);
                    }

                    lR2Entities.SaveChanges();
                }
            }

        }
    }
}

