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

namespace _4337Project
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _4337_d0h secondWindow = new _4337_d0h();

            // Show the window
            secondWindow.Show();
        }
        private void Button_Click1(object sender, RoutedEventArgs e)
        {
            _4337_SharifullnAnvar secondWindow = new _4337_SharifullnAnvar();

            // Show the window
            secondWindow.Show();
        }
        private void Button_Click3(object sender, RoutedEventArgs e)
        {
            _4337_Sergeev secondWindow = new _4337_Sergeev();

            // Show the window
            secondWindow.Show();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            _4337_hao secondWindow = new _4337_hao();

            // Show the window
            secondWindow.Show();
        }

        private void Vantsan_Click(object sender, RoutedEventArgs e)
        {
            _4337_VantsanMilena secondwindow = new _4337_VantsanMilena();
            secondwindow.Show();
        }
    }
}
