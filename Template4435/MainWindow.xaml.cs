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

        private void Maximov_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Максимов Роман Сергеевич", "4435_Максимов_Роман");
        }

        private void Adieva_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Адиева Айгуль Ринатовна", "4435_Адиева_Айгуль");
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
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
