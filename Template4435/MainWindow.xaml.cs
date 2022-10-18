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
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Data.Entity;
using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;

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
            MessageBox.Show("Автор: Сабиров Зульфат Зуфарович","4435_Сабиров_Зульфат");
        }
        private void toWindowBtn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Мартынов Максим Дмитриевич, 19 лет, группа_4435","4435_Мартынов");
        }
        private void AzatBtn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Хакимзянов Азат Гайсович", "4435_Хакимзянов_Азат");
        }
        private void BnnTask_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Назмутдинов Рузаль Ильгизович", "4435_Назмутдинов_Рузаль");
        }
        private void BtnCHELNY_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("ЕРКАШОВ 4435 19", "4435_ЕРКАШОВ");
        }
        private void BtnNikita_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("КРАВЧЕНКО 4435 16", "4435_КРАВЧЕНКО");
        }
        private void LR1_Shumilkin_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Шумилкин Александр Олегович", "4435_Шумилкин_Александр");
        }

        private void BtnCHELNYImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    DefaultExt = "*.xls;*.xlsx",
                    Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                    Title = "Выберите файл базы данных",
                    Multiselect = false
                };
                if (!(ofd.ShowDialog() == true))
                    return;

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(ofd.FileName);
                Excel.Worksheet ws = wb.Sheets[1];

                List<Clients> clients = new List<Clients>();
                for(int i = 2; i < ws.Rows.Count; i++)
                {
                    if (ws.Cells[i, "A"].Value != null)
                    {
                        Clients client = new Clients();
                        client.FIO = ws.Cells[i, "A"].Value;
                        client.id = Convert.ToInt32(ws.Cells[i, "B"].Value);
                        client.date_birth = Convert.ToDateTime(ws.Cells[i, "C"].Value);
                        client.adress_index = ws.Cells[i, "D"].Value.ToString();
                        client.adress_gorod = ws.Cells[i, "E"].Value.ToString().Substring(3);
                        client.adress_street = ws.Cells[i, "F"].Value;
                        client.adress_house = Convert.ToInt32(ws.Cells[i, "G"].Value.ToString());
                        client.adress_flat = Convert.ToInt32(ws.Cells[i, "H"].Value.ToString());
                        client.email = ws.Cells[i, "I"].Value;

                        clients.Add(client);

                    }
                    else break;
                }

                try
                {
                    using(Laba2ISRPOEntities context = new Laba2ISRPOEntities())
                    {
                        context.Clients.RemoveRange(context.Clients);
                        context.SaveChanges();
                        context.Clients.AddRange(clients);
                        context.SaveChanges();
                        MessageBox.Show("Данные импортированы");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.InnerException.InnerException.Message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnCHELNYExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<Clients> clients = new List<Clients>();
                using (Laba2ISRPOEntities context = new Laba2ISRPOEntities())
                {
                    clients =  context.Clients.ToList();
                }
                if(clients.Count > 0)
                {
                    Excel.Application app = new Excel.Application();
                    Excel.Workbook wb = app.Workbooks.Add();

                    Excel.Worksheet ws = (Excel.Worksheet)wb.Sheets.Add(After: wb.ActiveSheet);
                    ws.Name = "от 20 до 29";
                    Excel.Range rng = ws.get_Range("A1", "C1");
                    rng.Cells.Font.Bold = true;
                    rng.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    List<Clients> grouped = clients.Where(p => p.Age >= 20 && p.Age <= 29).ToList();
                    ws.Cells[1, "A"].Value = "Код клиента";
                    ws.Cells[1, "B"].Value = "ФИО";
                    ws.Cells[1, "C"].Value = "E-mail";
                    for (int i = 0; i < grouped.Count; i++)
                    {
                        ws.Cells[i+2, "A"].Value = grouped[i].id;
                        ws.Cells[i + 2, "B"].Value = grouped[i].FIO;
                        ws.Cells[i + 2, "C"].Value = grouped[i].email;
                    }
                    ws.Columns.EntireColumn.AutoFit();

                    ws = (Excel.Worksheet)wb.Sheets.Add(After: wb.ActiveSheet);
                    ws.Name = "от 30 до 39";
                    rng = ws.get_Range("A1", "C1");
                    rng.Cells.Font.Bold = true;
                    rng.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    grouped = clients.Where(p => p.Age >= 30 && p.Age <= 39).ToList();
                    ws.Cells[1, "A"].Value = "Код клиента";
                    ws.Cells[1, "B"].Value = "ФИО";
                    ws.Cells[1, "C"].Value = "E-mail";
                    for (int i = 0; i < grouped.Count; i++)
                    {
                        ws.Cells[i + 2, "A"].Value = grouped[i].id;
                        ws.Cells[i + 2, "B"].Value = grouped[i].FIO;
                        ws.Cells[i + 2, "C"].Value = grouped[i].email;
                    }
                    ws.Columns.EntireColumn.AutoFit();

                    ws = (Excel.Worksheet)wb.Sheets.Add(After: wb.ActiveSheet);
                    ws.Name = "от 40";
                    rng = ws.get_Range("A1", "C1");
                    rng.Cells.Font.Bold = true;
                    rng.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    grouped = clients.Where(p => p.Age >= 40).ToList();
                    ws.Cells[1, "A"].Value = "Код клиента";
                    ws.Cells[1, "B"].Value = "ФИО";
                    ws.Cells[1, "C"].Value = "E-mail";
                    for (int i = 0; i < grouped.Count; i++)
                    {
                        ws.Cells[i + 2, "A"].Value = grouped[i].id;
                        ws.Cells[i + 2, "B"].Value = grouped[i].FIO;
                        ws.Cells[i + 2, "C"].Value = grouped[i].email;
                    }
                    ws.Columns.EntireColumn.AutoFit();

                    ws = wb.Sheets[1];
                    ws.Delete();

                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
                    sfd.ShowDialog();
                    if (sfd.FileName != "")
                    {
                        wb.SaveAs(sfd.FileName);
                        wb.Close();
                        Process.Start(sfd.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnCHELNYExportJSON_Click(object sender, RoutedEventArgs e)
        {
            using (Laba2ISRPOEntities usersEntities = new  Laba2ISRPOEntities())
            {
                var allClients = usersEntities.Clients.ToList().OrderBy(s => s.Category).OrderBy(p=>p.FIO).ToList();
                var clientsCategories = allClients.GroupBy(s => s.Category).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                foreach (var group in clientsCategories) 
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = Convert.ToString(allClients.Where(g => g.Category == group.Key).FirstOrDefault().Category);
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table studentsTable =
                    document.Tables.Add(tableRange, group.Count() + 1, 3);
                    studentsTable.Borders.InsideLineStyle = studentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    studentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange = studentsTable.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = studentsTable.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = studentsTable.Cell(1, 3).Range;
                    cellRange.Text = "E-mail";
                    studentsTable.Rows[1].Range.Bold = 1;
                    studentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int i = 1;
                    foreach (var currentStudent in group)
                    {
                        cellRange = studentsTable.Cell(i + 1, 1).Range;
                        cellRange.Text = currentStudent.id.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(i + 1, 2).Range;
                        cellRange.Text = currentStudent.FIO;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = studentsTable.Cell(i + 1, 3).Range;
                        cellRange.Text = currentStudent.email;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        i++;
                    }
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }

                app.Visible = true;
            }
        }

        private void BtnCHELNYImportJSON_Click(object sender, RoutedEventArgs e)
        {
            using (Laba2ISRPOEntities entities = new Laba2ISRPOEntities())
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    DefaultExt = "*json",
                    Filter = "файл json|*.json",
                    Title = "Выберите файл json",
                    Multiselect = false
                };
                if (!(ofd.ShowDialog() == true))
                    return;
                try
                {
                    using (StreamReader r = new StreamReader(ofd.FileName))
                    {
                        entities.Clients.RemoveRange(entities.Clients);
                        entities.SaveChanges();
                        string json = r.ReadToEnd();
                        var options = new JsonSerializerOptions();
                        options.Converters.Add(new CustomDateTimeConverter());
                        List<Clients> items = JsonSerializer.Deserialize<List<Clients>>(json, options);
                        entities.Clients.AddRange(items);
                        entities.SaveChanges();
                        MessageBox.Show("Данные импортированы");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);
        static Process GetExcelProcess(Excel.Application excelApp)
        {
            GetWindowThreadProcessId(excelApp.Hwnd, out int id);
            return Process.GetProcessById(id);
        }

        private void ExcelImportBut_Click(object sender, RoutedEventArgs e)
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
            /*
            ObjWorkBook = null;
            ObjWorkExcel = null;
            ObjWorkSheet = null;
            lastCell = null;
            */
            Process process = GetExcelProcess(ObjWorkExcel);
            process.Kill();
            GC.Collect();
            try
            {
                using (isproBDEntities isproBD = new isproBDEntities())
                {
                    if (isproBD.Emloyee.FirstOrDefault() != null)
                    {
                         isproBD.Emloyee.RemoveRange(isproBD.Emloyee.ToList());
                         isproBD.SaveChanges();
                    }
                    for (int i = 1; i < _rows; i++)
                    {
                        isproBD.Emloyee.Add(new Emloyee() { CodeStaff = list[i, 0], Post = list[i, 1], FIO = list[i, 2], Login = list[i, 3], Password = list[i, 4], LastEnt = list[i, 5], Ent = list[i, 6] });
                    }
                        isproBD.SaveChanges();
                }
                MessageBox.Show("Данные успешно импортированы");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ExcelExportBut_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<Emloyee>> keyValues = new Dictionary<string, List<Emloyee>>(); 
            using (isproBDEntities isproBD = new isproBDEntities())
            {
                if (isproBD.Emloyee.FirstOrDefault() == null)
                {
                    MessageBox.Show("База данный пуста!");
                    return;
                }
                foreach(Emloyee em in isproBD.Emloyee)
                {
                    if (!keyValues.ContainsKey(em.Post))
                    {
                        keyValues.Add(em.Post, new List<Emloyee>() {em });
                    }
                    else
                    {
                        keyValues[em.Post].Add(em);
                    }
                }
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "файл Excel (Spisok.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == false)
                return;

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = keyValues.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < keyValues.Count(); i++)
            {
                string key = keyValues.Keys.ToArray()[i];
                Excel.Worksheet worksheet = app.Worksheets.Item[i+1];
                worksheet.Cells[1][1] = "Id";
                worksheet.Cells[2][1] = "Фио";
                worksheet.Cells[3][1] = "Логин";
                int j = 2;
                foreach (Emloyee emp in keyValues[key])
                {
                    worksheet.Cells[1][j] = emp.Id.ToString();
                    worksheet.Cells[2][j] = emp.FIO;
                    worksheet.Cells[3][j] = emp.Login;
                    j++;
                }
                worksheet.Columns.AutoFit();
                worksheet.Name = key;

            }

            if (saveFileDialog.FileName != "")
            {
                workbook.SaveAs(saveFileDialog.FileName);
                workbook.Close();
                Process.Start(saveFileDialog.FileName);
            }

            app.Quit();
        }

        private void JsonImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json|*.json",
                Title = "Выберите файл Json"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            using (isproBDEntities isproBD = new isproBDEntities())
            {
                List<Emloyee> listE;
                using (StreamReader r = new StreamReader(ofd.FileName))
                {
                    string s = r.ReadToEnd();
                    listE = JsonSerializer.Deserialize<List<Emloyee>>(s, new JsonSerializerOptions());
                }
                if (isproBD.Emloyee.FirstOrDefault() != null)
                {
                    isproBD.Emloyee.RemoveRange(isproBD.Emloyee.ToList());
                    isproBD.SaveChanges();
                }
                isproBD.Emloyee.AddRange(listE);
                isproBD.SaveChanges();
                MessageBox.Show("Данные успешно импортированы");
            }
        }

        private void ExportWord_Click(object sender, RoutedEventArgs e)
        {
            Dictionary<string, List<Emloyee>> keyValues = new Dictionary<string, List<Emloyee>>();
            using (isproBDEntities isproBD = new isproBDEntities())
            {
                if (isproBD.Emloyee.FirstOrDefault() == null)
                {
                    MessageBox.Show("База данный пуста!");
                    return;
                }
                foreach (Emloyee em in isproBD.Emloyee)
                {
                    if (!keyValues.ContainsKey(em.Post))
                    {
                        keyValues.Add(em.Post, new List<Emloyee>() { em });
                    }
                    else
                    {
                        keyValues[em.Post].Add(em);
                    }
                }
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "файл Word (Word.docx)|*.docx";
            if (saveFileDialog.ShowDialog() == false)
                return;

            var app = new Word.Application();
            Word.Document document = app.Documents.Add();

            Word.Paragraph paragraph = document.Paragraphs.Add();
            
            foreach(string key in keyValues.Keys)
            {
                //Заголовок
                Word.Paragraph Zagolovok = document.Paragraphs.Add();
                Zagolovok.Range.Text = key;
                Zagolovok.Range.InsertParagraphAfter();

                //Cоздание и форматирование таблицы
                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table EmployeeTable = document.Tables.Add(tableRange, keyValues[key].Count + 1, 3);
                EmployeeTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                EmployeeTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;


                //Название строк
                Word.Range cellRange;
                cellRange = EmployeeTable.Cell(1, 1).Range;
                cellRange.Text = "ID";
                cellRange = EmployeeTable.Cell(1, 2).Range;
                cellRange.Text = "ФИО";
                cellRange = EmployeeTable.Cell(1, 3).Range;
                cellRange.Text = "Логин";
                EmployeeTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                //Заполнение
                int i = 1;
                foreach (Emloyee CurEmloyee in keyValues[key])
                {
                    cellRange = EmployeeTable.Cell(i + 1, 1).Range;
                    cellRange.Text = CurEmloyee.Id.ToString();
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = EmployeeTable.Cell(i + 1, 2).Range;
                    cellRange.Text = CurEmloyee.FIO;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = EmployeeTable.Cell(i + 1, 3).Range;
                    cellRange.Text = CurEmloyee.Login;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    i++;
                }

                Word.Paragraph countEmployeeParagraph = document.Paragraphs.Add();
                Word.Range countStudentsRange = countEmployeeParagraph.Range;
                countStudentsRange.Text = $"Количество сотрудников по должности - {keyValues[key].Count()}";
                countStudentsRange.Font.Color = Word.WdColor.wdColorDarkRed;
                countStudentsRange.InsertParagraphAfter();
                document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            }

            if (saveFileDialog.FileName != "")
            {
                document.SaveAs(saveFileDialog.FileName);
                Process.Start(saveFileDialog.FileName);
            }
            app.Quit();
        }
    }

    public class CustomDateTimeConverter : JsonConverter<DateTime>
    {
        public override DateTime Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            return DateTime.ParseExact(reader.GetString(), "dd.MM.yyyy", null);
        }

        public override void Write(Utf8JsonWriter writer, DateTime value, JsonSerializerOptions options)
        {
            //Don't implement this unless you're going to use the custom converter for serialization too
            throw new NotImplementedException();
        }
    }
}
