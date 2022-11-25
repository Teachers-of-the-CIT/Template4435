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
using System.IO;
using System.Data.Entity;
using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

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
                    entities.Employee.Add(new Employee()
                    {
                        CodeStaff = list[i, 0],
                        Position = list[i, 1],
                        FullName = list[i, 2],
                        Log = list[i, 3],
                        Password = list[i, 4],
                        LastEnter = list[i, 5],
                        TypeEnter = list[i, 6]
                    });
                }
                entities.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Employee> allUsers;
            
            var allDoljnosti = _entities.Employee.Select(u => new {Doljnost = u.Position}).Distinct().ToList();
            allUsers = _entities.Employee.ToList().OrderBy(g => g.CodeStaff).ToList();
            
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
                var usersCategories = allUsers.GroupBy(s => s.Position).ToList();
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
                        foreach (Employee user in allUsers)
                        {
                            if (user.Position == users.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = user.CodeStaff;
                                worksheet.Cells[2][startRowIndex] = user.FullName;
                                worksheet.Cells[3][startRowIndex] = user.Log;
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

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<Employee>> keyValues = new Dictionary<string, List<Employee>>();
            using (Entities entities = new Entities())
            {
                if (entities.Employee.FirstOrDefault() == null)
                {
                    MessageBox.Show("База данных пуста");
                    return;
                }
                foreach (Employee em in entities.Employee)
                {
                    if (!keyValues.ContainsKey(em.Position))
                    {
                        keyValues.Add(em.Position, new List<Employee>() { em });
                    }
                    else
                    {
                        keyValues[em.Position].Add(em);
                    }
                }
            }

            var app = new Word.Application();
            Word.Document document = app.Documents.Add();

            Word.Paragraph paragraph = document.Paragraphs.Add();

            foreach (string key in keyValues.Keys)
            {
                //Заголовок
                Word.Paragraph Zagolovok = document.Paragraphs.Add();
                Zagolovok.Range.Text = key;
                Zagolovok.set_Style("Заголовок 1");
                Zagolovok.Range.InsertParagraphAfter();

                //Cоздание и форматирование таблицы
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table EmployeeTable = document.Tables.Add(tableRange, keyValues[key].Count + 1, 3);
                EmployeeTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                EmployeeTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                EmployeeTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //Название строк
                Word.Range cellRange;
                cellRange = EmployeeTable.Cell(1, 1).Range;
                cellRange.Text = "Код сотрудника";
                cellRange = EmployeeTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = EmployeeTable.Cell(1, 3).Range;
                cellRange.Text = "Логин";
                EmployeeTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                EmployeeTable.Rows[1].Range.Bold = 1;

                //Заполнение
                int i = 1;
                foreach (Employee CurEmloyee in keyValues[key])
                {
                    cellRange = EmployeeTable.Cell(i + 1, 1).Range;
                    cellRange.Text = CurEmloyee.CodeStaff.ToString();
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = EmployeeTable.Cell(i + 1, 2).Range;
                    cellRange.Text = CurEmloyee.FullName;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = EmployeeTable.Cell(i + 1, 3).Range;
                    cellRange.Text = CurEmloyee.Log;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    i++;
                }

                Word.Paragraph countEmployeeParagraph = document.Paragraphs.Add();
                Word.Range countStudentsRange = countEmployeeParagraph.Range;
                countStudentsRange.Text = $"Количество сотрудников данной должности - {keyValues[key].Count()}";
                countStudentsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                countStudentsRange.InsertParagraphAfter();
                document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                app.Visible = true;
                document.SaveAs2(@"D:\outputFileWord.docx");
                document.SaveAs2(@"D:\outputFilePdf.pdf",
                Word.WdExportFormat.wdExportFormatPDF);
            }
        }

        private void BnJSONImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json|*.json",
                Title = "Выберите файл Json"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            using (Entities entities = new Entities())
            {
                List<Employee> listE;
                using (StreamReader r = new StreamReader(ofd.FileName))
                {
                    string s = r.ReadToEnd();
                    listE = JsonSerializer.Deserialize<List<Employee>>(s, new JsonSerializerOptions());
                }
                if (entities.Employee.FirstOrDefault() != null)
                {
                    entities.Employee.RemoveRange(entities.Employee.ToList());
                    entities.SaveChanges();
                }
                entities.Employee.AddRange(listE);
                entities.SaveChanges();
            }
        }
    }
    }

