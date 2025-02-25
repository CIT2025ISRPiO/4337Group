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

namespace _4337Project
{
    /// <summary>
    /// Логика взаимодействия для _4337_Vafin.xaml
    /// </summary>
    public partial class _4337_Vafin : Window
    {
        public _4337_Vafin()
        {
            InitializeComponent();
        }
        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                string connectionString = "Server=DESKTOP-3161DTA;Database=LabaISRPO;User Id=your_username;Integrated Security=True;";
                string tableName = "Orders";

                Vafin.ImportData(filePath, connectionString, tableName);

                MessageBox.Show("Данные успешно импортированы!");
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string outputFilePath = saveFileDialog.FileName;
                string connectionString = "Server=DESKTOP-3161DTA;Database=LabaISRPO;User Id=your_username;Integrated Security=True;";
                string tableName = "Orders";

                Vafin.ExportData(connectionString, tableName, outputFilePath);

                MessageBox.Show("Данные успешно экспортированы!");
            }
        }
    }
}
