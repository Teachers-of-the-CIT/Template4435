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
using Microsoft.Office.Interop;


namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для Rahimov.xaml
    /// </summary>
    public partial class Rahimov : System.Windows.Window
    {
        public static DateTime JavaTimeStampToDateTime(string javaTimeStamp)
        {
            // Java timestamp is milliseconds past epoch
            DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
            dateTime = dateTime.AddMilliseconds(Convert.ToDouble(javaTimeStamp)).ToLocalTime();
            return dateTime;
        }
        public Rahimov()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(open.ShowDialog() == true))
                return;
            string[,] list;
            Microsoft.Office.Interop.Excel.Application ObjWorkExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(open.FileName);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using(var db = new ISRPO2Entities())
            {
                for (int i = 1; i < _rows; i++)
                {
                    int nullCollumns = 0;
                    for (int j = 0; j < _columns; j++)
                    {
                        if (String.IsNullOrEmpty(list[i, j]))
                            nullCollumns++;
                    }
                    if (nullCollumns == _columns)
                        continue;
                    db.Orders.Add(new Orders()
                    {
                        ID = Convert.ToInt32(list[i, 0]),
                        Kod = list[i, 1],
                        DateOfCreating = list[i, 2],
                        TimeOfCreating = list[i, 3],
                        KodKlient = Convert.ToInt32(list[i, 4]),
                        Service = list[i, 5],
                        Status = list[i, 6],
                        DateOfClosing = list[i, 7],
                        TImeOfRental = list[i, 8]
                    });

                }
                db.SaveChanges();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

            List<Orders> allOrders;
            List<string> allDate;
            using (var db = new ISRPO2Entities())
            {
                allOrders = (from o in db.Orders
                             orderby o.Kod
                             select o).ToList();
                allDate = (from d in allOrders
                           group d by d.DateOfCreating into g
                           select g.Key).ToList();
                
            }
            var app = new Microsoft.Office.Interop.Excel.Application();
            app.SheetsInNewWorkbook = allDate.Count;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            app.Visible = true;
            for (int i = 0; i < allDate.Count; i++)
            {
                int startRowIndex = 1;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = allDate[i];
                worksheet.Cells[1][startRowIndex] = "Id";
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[3][startRowIndex] = "Код клиента";
                worksheet.Cells[4][startRowIndex] = "Услуги";
                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1],
                    worksheet.Cells[4][1]];
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;
                startRowIndex++;
                foreach(Orders order in allOrders)
                {
                    if (order.DateOfCreating != allDate[i])
                        continue;
                    worksheet.Cells[1][startRowIndex] = order.ID;
                    worksheet.Cells[2][startRowIndex] = order.Kod;
                    worksheet.Cells[3][startRowIndex] = order.KodKlient;
                    worksheet.Cells[4][startRowIndex] = order.Service;
                    startRowIndex++;
                }
                Microsoft.Office.Interop.Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[4][startRowIndex - 1]];
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = 
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = 
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = 
                Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            //app.Visible = true;
        }
    }
}
