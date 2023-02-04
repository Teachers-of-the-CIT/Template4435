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

        private void Aigulyash_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Хайруллина Айгуль Рафисовна, 19 лет", "4435_Хайруллина_Айгуль");
        }
        
        private void Chernovskiy_4435_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Черновский Даниэль", "Черновский Даниэль 4435", MessageBoxButton.OK);
        }
        
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Зиганшина Зиля Рушановна", "4435_Зиганшина_Зиля");
        }
        
        private void Rahimov_Click(object sender, RoutedEventArgs e)
        {
            new Rahimov().Show();
        }
    }
}
