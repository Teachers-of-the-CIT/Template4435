using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BnTask_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Сабиров Зульфат Зуфарович", "4435_Сабиров_Зульфат");
        }
        private void AzatBtn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Хакимзянов Азат Гайсович", "4435_Хакимзянов_Азат");
        }
        private void BnnTask_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Назмутдинов Рузаль Ильгизович", "4435_Назмутдинов_Рузаль");
        }

        private void ImportB_Click(object sender, RoutedEventArgs e)
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
            using (ISEntities2 sd = new ISEntities2())
            {
                for (int i = 0; i < _rows; i++)
                {
                    sd.User.Add(new User()
                    {
                        FullName = list[i, 0],
                        CodeClient = list[i, 1],
                        BirthDate = list[i, 2],
                        Index = list[i, 3],
                        City = list[i, 4],
                        Street = list[i, 5],
                        Home = list[i, 6],
                        Kvartira = list[i, 7],
                        E_mail = list[i, 8]
                    });
                }
                sd.SaveChanges();
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int cat1 = 0;
            int cat2 = 0;
            int cat3 = 0;
            List<User> users;
            using (ISEntities2 usersEntities = new ISEntities2())
            {
                users = usersEntities.User.ToList();
            }
            var ageCategories = users.OrderBy(o => o.BirthDate).GroupBy(s => s.Id)
                    .ToDictionary(g => g.Key, g => g.Select(s => new { s.Id, s.FullName, s.BirthDate, s.E_mail }).ToArray());
            var app = new Word.Application();
            foreach (var item in users)
            {
                string a = item.BirthDate;
                string g = a.Substring(a.Length - 4);
                int b = 2022 - Int32.Parse(g);
                if(b >= 20 && b <= 29)
                {
                    cat1++;
                }
                else if (b >= 30 && b <= 39)
                {
                    cat2++;
                }
                else if (b >= 40)
                {
                    cat3++;
                }
            }
            Word.Document document = app.Documents.Add();
            for (int i = 0; i < 1; i++)
            {
                var data = users;
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = $"Категория {i + 1}";
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var userTable = document.Tables.Add(tableRange, cat1 + 1, 4);
                userTable.Borders.InsideLineStyle = userTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                userTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;
                cellRange = userTable.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = userTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = userTable.Cell(1, 3).Range;
                cellRange.Text = "Email";
                cellRange = userTable.Cell(1, 4).Range;
                cellRange.Text = "Возраст";
                userTable.Rows[1].Range.Bold = 1;
                userTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                int row = 1;
                var stepSize = 1;
                foreach (var group in data)
                {
                    string a = group.BirthDate;
                    string g = a.Substring(a.Length - 4);
                    int b = 2022 - Int32.Parse(g);
                    if (b >= 20 && b <= 29)
                    {
                        cellRange = userTable.Cell(row + stepSize, 1).Range;
                        cellRange.Text = group.CodeClient.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 2).Range;
                        cellRange.Text = group.FullName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 3).Range;
                        cellRange.Text = group.E_mail.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 4).Range;
                        cellRange.Text = b.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        row++;                       

                    }
                    }
                Word.Paragraph countCostsParagraph = document.Paragraphs.Add();
                Word.Range countCostsRange = countCostsParagraph.Range;
                countCostsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                countCostsRange.InsertParagraphAfter();
                document.Words.Last.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
            for (int l = 0; l < 1; l++)
            {
                var data = users;
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = $"Категория {l + 2}";
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var userTable = document.Tables.Add(tableRange, cat2 + 1, 4);
                userTable.Borders.InsideLineStyle = userTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                userTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;
                cellRange = userTable.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = userTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = userTable.Cell(1, 3).Range;
                cellRange.Text = "Email";
                cellRange = userTable.Cell(1, 4).Range;
                cellRange.Text = "Возраст";
                userTable.Rows[1].Range.Bold = 1;
                userTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                int row = 1;
                var stepSize = 1;
                foreach (var group in data)
                {
                    string a = group.BirthDate;
                    string g = a.Substring(a.Length - 4);
                    int b = 2022 - Int32.Parse(g);
                    if (b >= 30 && b <= 39)
                    {
                        cellRange = userTable.Cell(row + stepSize, 1).Range;
                        cellRange.Text = group.CodeClient.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 2).Range;
                        cellRange.Text = group.FullName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 3).Range;
                        cellRange.Text = group.E_mail.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 4).Range;
                        cellRange.Text = b.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        row++;

                    }
                }
                Word.Paragraph countCostsParagraph = document.Paragraphs.Add();
                Word.Range countCostsRange = countCostsParagraph.Range;
                countCostsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                countCostsRange.InsertParagraphAfter();
                document.Words.Last.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
            for (int m = 0; m < 1; m++)
            {
                var data = users;
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = $"Категория {m + 3}";
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                var userTable = document.Tables.Add(tableRange, cat3 + 1, 4);
                userTable.Borders.InsideLineStyle = userTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                userTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                Word.Range cellRange;
                cellRange = userTable.Cell(1, 1).Range;
                cellRange.Text = "Код клиента";
                cellRange = userTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = userTable.Cell(1, 3).Range;
                cellRange.Text = "Email";
                cellRange = userTable.Cell(1, 4).Range;
                cellRange.Text = "Возраст";
                userTable.Rows[1].Range.Bold = 1;
                userTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                int row = 1;
                var stepSize = 1;
                foreach (var group in data)
                {
                    string a = group.BirthDate;
                    string g = a.Substring(a.Length - 4);
                    int b = 2022 - Int32.Parse(g);
                    if (b >= 40)
                    {
                        cellRange = userTable.Cell(row + stepSize, 1).Range;
                        cellRange.Text = group.CodeClient.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 2).Range;
                        cellRange.Text = group.FullName;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 3).Range;
                        cellRange.Text = group.E_mail.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = userTable.Cell(row + stepSize, 4).Range;
                        cellRange.Text = b.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        row++;

                    }
                }
                Word.Paragraph countCostsParagraph = document.Paragraphs.Add();
                Word.Range countCostsRange = countCostsParagraph.Range;
                countCostsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                countCostsRange.InsertParagraphAfter();
                document.Words.Last.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
            app.Visible = true;
            document.SaveAs2(@"D:\outputFileWord.docx");
            document.SaveAs2(@"D:\outputFilePdf.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }
    

        private void ExportB_Click(object sender, RoutedEventArgs e)
        {
            List<User> allUser;
            using (ISEntities2 lR2Entities = new ISEntities2())
            {
                allUser = lR2Entities.User.ToList();
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == false)
                return;
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < 3; i++)
            {
                int startRowIndex = 1;
                int str = i;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = "Категория "+ (str+1).ToString();
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                worksheet.Cells[3][startRowIndex] = "Возраст";
                worksheet.Cells[4][startRowIndex] = "E-mail";
                

                startRowIndex++;
                int category = 0;
                int j = 2;
                foreach (var user in allUser)
                {

                    if (j!=0 && j<=70 && user.BirthDate !="")
                    {
                        string a = user.BirthDate;
                        string g = a.Substring(a.Length - 4);
                        int  b = 2022-Int32.Parse(g);
                        startRowIndex++;
                        
                        if(b>=20 && b<=29 && str==0)
                        {
                            worksheet.Cells[1][j] = user.CodeClient;
                            worksheet.Cells[2][j] = user.FullName;
                            worksheet.Cells[3][j] = b;
                            worksheet.Cells[4][j] = user.E_mail;
                            j++;
                            continue;
                        }
                        else if(b >= 30 && b <= 39 && str==1)
                        {
                            
                            worksheet.Cells[1][j] = user.CodeClient;
                            worksheet.Cells[2][j] = user.FullName;
                            worksheet.Cells[3][j] = b;
                            worksheet.Cells[4][j] = user.E_mail;
                            j++;
                            
                        }
                        else if (b >= 40 && str == 2)
                        {

                            worksheet.Cells[1][j] = user.CodeClient;
                            worksheet.Cells[2][j] = user.FullName;
                            worksheet.Cells[3][j] = b;
                            worksheet.Cells[4][j] = user.E_mail;
                            j++;

                        }

                    }
                    else if (j>70)
                    {
                        continue;
                        
                    }
                       


                }
                
                worksheet.Columns.AutoFit();

            }

            if (saveFileDialog.FileName != "")
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                    workbook.Close();
                    Process.Start(saveFileDialog.FileName);
                }
                app.Quit();
        }


        public async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Title = "Выберите файл JSON"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate);
            List<User> users = await JsonSerializer.DeserializeAsync<List<User>>(fs);
            using (ISEntities2 lR2Entities = new ISEntities2())
            {
                foreach (User user in users)
                {
                    User userr = new User();
                    userr.Id = user.Id;
                    userr.BirthDate = user.BirthDate;
                    userr.CodeClient = user.CodeClient;
                    userr.E_mail = user.E_mail;
                    userr.City = user.City;
                    userr.Home =  user.Home;
                    userr.Index= user.Index;
                    userr.FullName = user.FullName;
                    userr.Street = user.Street;
                    userr.Kvartira = user.Kvartira;
                    lR2Entities.User.Add(userr);
                }
                lR2Entities.SaveChanges();
            }
        }

        
    }
}
