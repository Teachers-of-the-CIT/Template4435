using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity;
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
using Template4435.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для MartynovWindow.xaml
    /// </summary>
    public partial class MartynovWindow : Window
    {
        public MartynovWindow()
        {
            InitializeComponent();
        }

        private void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (5.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list; 
            Excel.Application objWorkExcel = new Excel.Application();
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
            var lastCell = objWorkExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int columns = (int)lastCell.Column;
            int rows = (int)lastCell.Row;
            list = new string[rows, columns];
            for (int j = 0; j < columns; j++)
                for (int i = 0; i < rows; i++)
                    list[i, j] = objWorkExcel.Cells[i + 1, j + 1].Text;
            objWorkBook.Close(false, Type.Missing, Type.Missing);
            objWorkExcel.Quit();
            GC.Collect();

            for (int i = 1; i < rows; i++)
            {
                UsersEntities.GetContext().Users.Add(new User() { Role = list[i, 0],  FIO = list[i, 1], Login = list[i, 2], Password = list[i, 3]});
            }
            UsersEntities.GetContext().SaveChanges();
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            var role = UsersEntities.GetContext().Users.ToList().OrderBy(u => u.Role).ToList();
        }
    }
}
