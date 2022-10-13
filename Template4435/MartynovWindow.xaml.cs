using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Security.Cryptography;
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
                Filter = "файлы Excel |*.xlsx",
                Title = "Выберите Excel файл"
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
                UsersEntities.GetContext().Users
                    .Add(new User() { Role = list[i, 0],  FIO = list[i, 1], Login = list[i, 2], Password = list[i, 3]});
            }
            UsersEntities.GetContext().SaveChanges();
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            var roles = UsersEntities.GetContext().Users
                .ToList()
                .GroupBy(r => r.Role)
                .ToList();
            var users = UsersEntities.GetContext().Users
                .ToList()
                .OrderBy(u => u.Role)
                .ToList();

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = roles.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < roles.Count(); i++)
            {
                int startRowIndex = 2;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = roles[i].Key.ToString();
                worksheet.Cells[1][startRowIndex] = "Логин";
                worksheet.Cells[2][startRowIndex] = "Пароль";
                startRowIndex++;
                foreach (var user in users)
                {
                    if (roles[i].Key == user.Role)
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = roles[i].Key;
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Bold = true;
                        worksheet.Cells[1][startRowIndex] = user.Login;
                        worksheet.Cells[2][startRowIndex] = GetHash(user.Password);
                        startRowIndex++;
                    }
                    else
                    {
                        continue;
                    }

                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 1]];
                    rangeBorders.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    worksheet.Columns.AutoFit();
                }
                app.Visible = true;
            }  
        }

        public string GetHash(string input)
        {
            var md5 = MD5.Create();
            var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));

            return Convert.ToBase64String(hash);
        }
    }
}
