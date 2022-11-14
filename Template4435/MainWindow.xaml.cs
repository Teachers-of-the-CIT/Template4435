using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Btn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Багаутинова Софья Вахтанговна", "4435_Багаутинова_Софья");
        }

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
            }
        }

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

    }
}

