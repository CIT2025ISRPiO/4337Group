using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using ExcelDataReader;
using System.Data;



namespace _4337Project
{
    /// <summary>
    /// Interaction logic for d0h_ImportWindow.xaml
    /// </summary>
    public partial class d0h_ImportWindow : Window
    {
        private string connectionString = @"Data Source=DESKTOP-AF0FDGA;Initial Catalog=ISRPO_db;Integrated Security=True;";
        public d0h_ImportWindow()
        {
            InitializeComponent();
            AllowDrop = true;
        }

        private void Window_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();

                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка подключения: {ex.Message}");
                }
            }

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string excelFilePath = files[0];
                    ImportExcel(excelFilePath);
                }
            }
        }


        private void ImportExcel(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });

                        DataTable dataTable = result.Tables[0]; 

                        foreach (DataRow row in dataTable.Rows)
                        {
                            if (!int.TryParse(row["Код клиента"].ToString(), out int clientCode))
                            {
                                row["Код клиента"] = 0;
                            }

                        }

                        // Теперь безопасно вставляем данные в базу
                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
                            {
                                bulkCopy.DestinationTableName = "Clients";
                                bulkCopy.ColumnMappings.Add("ФИО", "FIO");
                                bulkCopy.ColumnMappings.Add("Код клиента", "ClientCode");
                                bulkCopy.ColumnMappings.Add("Дата рождения", "DateOfBirth");
                                bulkCopy.ColumnMappings.Add("Индекс", "Index");
                                bulkCopy.ColumnMappings.Add("Город", "City");
                                bulkCopy.ColumnMappings.Add("Улица", "Street");
                                bulkCopy.ColumnMappings.Add("Дом", "House");
                                bulkCopy.ColumnMappings.Add("Квартира", "Flat");
                                bulkCopy.ColumnMappings.Add("E-mail", "Email");

                                bulkCopy.WriteToServer(dataTable);
                            }
                        }
                    }
                }

                MessageBox.Show("Импорт завершен!");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка импорта: {ex.Message}");
            }
        }


    }
}
