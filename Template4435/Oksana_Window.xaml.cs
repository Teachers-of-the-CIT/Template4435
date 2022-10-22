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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для Oksana_Window.xaml
    /// </summary>
    public partial class Oksana_Window : Window
    {
        public Oksana_Window()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, RoutedEventArgs e)
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
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();


            using (ISRPOEntities ISRPO = new ISRPOEntities())
            {
                for (int i = 0; i < _rows; i++)
                {
                    ISRPO.Users.Add(new Users()
                    {
                        Position =
                    list[i, 1],
                        FIO = list[i, 2],
                        Login = list[i, 3],
                        Pass = list[i, 4],
                        Last_ent = list[i, 5],
                        Type_ent = list[i, 6]
                    });
                }
                ISRPO.SaveChanges();
            }

        }
        public List<Users> allUsers;
        private ISRPOEntities ISRPO = new ISRPOEntities();
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            allUsers = ISRPO.Users.ToList();

            var list_types = allUsers.Select(x => x.Type_ent).Distinct().ToList();//лист только с типами входа 
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = list_types.Count();//сколько типов входа - столько листов 
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < list_types.Count(); i++)
            {
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(list_types[i]);//имя листа - тип входа 
                int startRowIndex = 1;//счетчик строк 
                worksheet.Cells[1][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex] = "Должность";
                worksheet.Cells[3][startRowIndex] = "Логин";
                Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;

                startRowIndex++;
                foreach (var user in allUsers)
                {
                    if (user.Type_ent == worksheet.Name)
                    {
                        worksheet.Cells[1][startRowIndex] = user.Id;
                        worksheet.Cells[2][startRowIndex] = user.Position;
                        worksheet.Cells[3][startRowIndex] = user.Login;
                        startRowIndex++;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex - 1]];
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
    }
}
    

