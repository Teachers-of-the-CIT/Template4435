using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Markup;
using System.Xaml;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static EmployeeEntities3 _context = new EmployeeEntities3();

        private void Load()
        {
            Employee.ItemsSource = _context.Position.ToList();
        }

        public MainWindow()
        {
            InitializeComponent();
            Load();
        }        

        private void Btn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Багаутинова Софья Вахтанговна", "4435_Багаутинова_Софья");
        }
        /// <summary>
        /// Импорт данных из Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list; // массив для хранения данных из xlsx-файла
            Excel.Application ObjWorkExcel = new Excel.Application(); // экземпляр класса для работы с библитекй Introp
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName); // экземпляр класса для загрузки документа формата xlsx для работы с электронными таблицами
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //выбор xlsx-файла для дальнейшей работы с ним
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell); // определение последней ячейки таблицы
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 1; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (EmployeeEntities3 employeeEntities = new EmployeeEntities3())
            {
                for (int i = 1; i < _rows; i++)
                {
                    employeeEntities.Position.Add(new Position()
                    {
                        Code_Employee = list[i, 0],
                        Positions = list[i, 1],
                        FullName = list[i, 2],
                        Employee_Login = list[i, 3],
                        Employee_Password = list[i, 4],
                        Last_Entrance = list[i, 5],
                        Type_Of_Entrance = list[i, 6]
                    });
                }
                employeeEntities.SaveChanges();
                MessageBox.Show("Данные удачно импортировались!");
                Load();
            }
        }

        /// <summary>
        /// Экспорт данных в Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Position> positions = new List<Position>();

            using (EmployeeEntities3 employeeEntities = new EmployeeEntities3())
            {
                //выоборка на разделение по должности 
                positions = employeeEntities.Position.GroupBy(x => x.Positions).Select(x => x.FirstOrDefault()).ToList();

                //создание Ecxel
                var app = new Excel.Application();
                app.SheetsInNewWorkbook = positions.Count();
                Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

                for (int i = 0; i < positions.Count(); i++)
                {
                    int startRowIndex = 1;
                    // именование листов в Excel
                    Excel.Worksheet worksheet = app.Worksheets.Item[i + 1]; //создание нового листа
                    worksheet.Name = Convert.ToString(positions[i].Positions); //название листа согласно должности

                    //названия колонок
                    worksheet.Cells[1][1] = "Код клиента";
                    worksheet.Cells[2][1] = "ФИО";
                    worksheet.Cells[3][1] = "Логин";

                    startRowIndex++;

                    foreach (var position in employeeEntities.Position)
                    {
                        if (positions[i].Positions == position.Positions)
                        {
                            worksheet.Cells[1][startRowIndex] = position.Code_Employee;
                            worksheet.Cells[2][startRowIndex] = position.FullName;
                            worksheet.Cells[3][startRowIndex] = position.Employee_Login;
                            startRowIndex++;
                        }
                    }
                }
                app.Visible = true;
            }
        }

        /// <summary>
        /// Импорт JSON данных
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportJSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json (Spisok.json)|*.json",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            string json = File.ReadAllText(ofd.FileName); //чтение данных из json-файла
            var position = JsonSerializer.Deserialize<List<Position>>(json); //десереализация

            using (EmployeeEntities3 db = new EmployeeEntities3())
            {
                foreach (Position positions in position)
                {
                    Position position1 = new Position();
                    position1.Code_Employee = positions.Code_Employee;
                    position1.Positions = positions.Positions;
                    position1.FullName = positions.FullName;
                    position1.Employee_Login = positions.Employee_Login;
                    position1.Employee_Password = positions.Employee_Password;
                    position1.Last_Entrance = positions.Last_Entrance;
                    position1.Type_Of_Entrance = positions.Type_Of_Entrance;
                    db.Position.Add(position1);
                }
                db.SaveChanges();
                MessageBox.Show("Данные успешно импортировались!");
                Load();
            }
        }

        /// <summary>
        /// Экспорт JSON данных в Word
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportJSON_Click(object sender, RoutedEventArgs e)
        {
            // список для хранения отсортированных по ФИО данных
            var sortedByPosition = new List<Position>();
            // список названий всех улиц (будущие категории)
            var allPositionNames = new List<string>();

            // заполнение списка информацией из бд
            using (EmployeeEntities3 db = new EmployeeEntities3())
            {
                sortedByPosition = db.Position.ToList().OrderBy(x => x.Positions).ToList();

                foreach (var item in sortedByPosition)
                {
                    if (item.Positions != null && item.Positions != "")
                        allPositionNames.Add(item.Positions);
                }

                var distinctEntryType = allPositionNames.Distinct().ToList();

                var app = new Word.Application();
                var document = app.Documents.Add();
                var index = 0;

                for (var i = 0; i < distinctEntryType.Count(); i++)
                {
                    int rowCounter = 0;

                    foreach (var item in sortedByPosition)
                    {
                        if (item.Positions.Replace(" ", "") == distinctEntryType[index].Replace(" ", ""))
                            rowCounter++;
                    }

                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = distinctEntryType[index];
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var positionCategories = document.Tables.Add(tableRange, rowCounter + 1, 3);

                    positionCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    positionCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    positionCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = positionCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код клиента";
                    cellRange = positionCategories.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = positionCategories.Cell(1, 3).Range;
                    cellRange.Text = "Логин";

                    positionCategories.Rows[1].Range.Bold = 1;
                    positionCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;

                    foreach (var item in sortedByPosition)
                    {
                        if (item.Positions == distinctEntryType[index])
                        {
                            cellRange = positionCategories.Cell(count + 1, 1).Range;
                            cellRange.Text = item.Code_Employee;
                            cellRange = positionCategories.Cell(count + 1, 2).Range;
                            cellRange.Text = item.FullName;
                            cellRange = positionCategories.Cell(count + 1, 3).Range;
                            cellRange.Text = item.Employee_Login;
                            count++;
                        }
                    }
                    index++;
                    Word.Paragraph countCostsParagraph = document.Paragraphs.Add();
                    Word.Range countCostsRange = countCostsParagraph.Range;
                    countCostsRange.InsertParagraphAfter();
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                }
                for (int i = 0; i < 1; i++)
                {
                    var rowCounter = 0;
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Количество работников в системе.";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var streetCategories = document.Tables.Add(tableRange,4, 2);

                    streetCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    streetCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    streetCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = streetCategories.Cell(1, 1).Range;
                    cellRange.Text = "Должность";
                    cellRange = streetCategories.Cell(1, 2).Range;
                    cellRange.Text = "Количество";

                    streetCategories.Rows[1].Range.Bold = 1;
                    streetCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                            var admin = db.Position.Where(x => x.Positions == "Администратор").ToList();
                            var countAdmin = admin.Count();
                            cellRange = streetCategories.Cell(1 + 1, 1).Range;
                            cellRange.Text = "Администратор";
                            cellRange = streetCategories.Cell(1 + 1, 2).Range;
                            cellRange.Text = countAdmin.ToString();
                    rowCounter++;

                            var salesman = db.Position.Where(x => x.Positions == "Продавец").ToList();
                            var countSalesman = salesman.Count();
                            cellRange = streetCategories.Cell(2 + 1, 1).Range;
                            cellRange.Text = "Продавец";
                            cellRange = streetCategories.Cell(2 + 1, 2).Range;
                            cellRange.Text = countSalesman.ToString();
                    rowCounter++;

                            var older = db.Position.Where(x => x.Positions == "Старший смены").ToList();
                            var countOlder = older.Count();
                            cellRange = streetCategories.Cell(3 + 1, 1).Range;
                            cellRange.Text = "Старший смены";
                            cellRange = streetCategories.Cell(3 + 1, 2).Range;
                            cellRange.Text = countOlder.ToString();
                    rowCounter++;
                            count++;  
                }
                // показываем готовую книгу Excel
                app.Visible = true;
                document.SaveAs(@"C:\Users\polus\OneDrive\Рабочий стол\Соня\ИСРПО\outputFileWord.docx");
                document.SaveAs(@"C:\Users\polus\OneDrive\Рабочий стол\Соня\ИСРПО\outputFilePdf.pdf", Word.WdExportFormat.wdExportFormatPDF);
            }
        }
        

        private void Employee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}

