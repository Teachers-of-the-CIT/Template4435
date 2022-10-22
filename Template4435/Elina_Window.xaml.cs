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
    /// Логика взаимодействия для Elina_Window.xaml
    /// </summary>
    public partial class Elina_Window : Window
    {
        public Elina_Window()
        {
            InitializeComponent();
        }

        private void btn_import_Click(object sender, RoutedEventArgs e)
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
                    ISRPO.User.Add(new User()
                    {
                        pos = list[i, 1],
                        fio = list[i, 2],
                        login = list[i, 3],
                        pass = list[i, 4],
                        lastEnt = list[i, 5],
                        typeEnt = list[i, 6]
                    });
                }
                ISRPO.SaveChanges();
            }

        }
        public List<User> UserL;
        private ISRPOEntities ISRPO = new ISRPOEntities();
        private void btn_export_Click(object sender, RoutedEventArgs e)
        {
            UserL = ISRPO.User.ToList();

            var list_types = UserL.Select(x => x.typeEnt).Distinct().ToList();//лист только с типами входа 
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
                foreach (var user in UserL)
                {
                    if (user.typeEnt == worksheet.Name)
                    {
                        worksheet.Cells[1][startRowIndex] = user.id;
                        worksheet.Cells[2][startRowIndex] = user.pos;
                        worksheet.Cells[3][startRowIndex] = user.login;
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
