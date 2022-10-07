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
//using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Data.Entity;
using System.Diagnostics;
using System.Text.Json;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
//using Newtonsoft.Json.Linq;

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
                        client.adress_index = Convert.ToInt32(ws.Cells[i, "D"].Value);
                        client.adress_gorod = ws.Cells[i, "E"].Value.ToString().Substring(3);
                        client.adress_street = ws.Cells[i, "F"].Value;
                        client.adress_house = ws.Cells[i, "G"].Value.ToString();
                        client.adress_flat = ws.Cells[i, "H"].Value.ToString();
                        client.email = ws.Cells[i, "I"].Value;

                        clients.Add(client);

                    }
                    else break;
                }

                try
                {
                    using(Laba2ISRPOEntities context = new Laba2ISRPOEntities())
                    {
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
                        string json = r.ReadToEnd();
                        List<ClientsJSON> items = JsonConvert.DeserializeObject<List<ClientsJSON>>(json, new IsoDateTimeConverter { DateTimeFormat = "dd.MM.yyyy" });
                        entities.ClientsJSON.AddRange(items);
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
    }
}
