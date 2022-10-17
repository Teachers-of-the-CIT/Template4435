using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Template4435.Model;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;

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
                Filter = "Excel files |*.xlsx| All files(*.*)|*.*",
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

        private void ImportJSONBtn_Click(object sender, RoutedEventArgs e)
        {
            var users = new List<User>();
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json;*.json",
                Filter = "JSON files |*.json| All files(*.*)|*.*",
                Title = "Выберите JSON файл"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            using (FileStream fs = new FileStream($"{ofd.FileName}", FileMode.Open))
            {
                users = JsonSerializer.Deserialize<List<User>>(fs);
            }
            foreach (var user in users)
            {
                UsersEntities.GetContext().Users.Add(user);
            }
            UsersEntities.GetContext().SaveChanges();
        }

        private void ExportWordBtn_Click(object sender, RoutedEventArgs e)
        {
            var roles = UsersEntities.GetContext().Users
               .ToList()
               .GroupBy(r => r.Role)
               .ToList();
            var users = UsersEntities.GetContext().Users
                .ToList()
                .OrderBy(u => u.Role)
                .ToList();
            var app = new Word.Application();
            Word.Document document = app.Documents.Add();

            foreach (var role in roles)
            {
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = Convert.ToString(role.Key + "ы");
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table usersTable = document.Tables.Add(tableRange,  UsersEntities.GetContext().Users.ToList().Where(x => x.Role == role.Key).Count() + 1, 2);
                usersTable.Borders.InsideLineStyle = usersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range cellRange = usersTable.Cell(1, 1).Range;
                cellRange.Text = "Логин";
                cellRange = usersTable.Cell(1, 2).Range;
                cellRange.Text = "Пароль";

                int i = 1;
                foreach (var user in users)
                {
                    if (role.Key == user.Role)
                    {
                        cellRange = usersTable.Cell(i + 1, 1).Range;
                        cellRange.Text = user.Login.ToString();
                        cellRange = usersTable.Cell(i + 1, 2).Range;
                        cellRange.Text = GetHash(user.Password);
                        i++;
                    }
                   
                }
                usersTable.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                usersTable.Rows[1].Range.Bold = 1;
            }
            app.Visible = true;
        }
    }
}
