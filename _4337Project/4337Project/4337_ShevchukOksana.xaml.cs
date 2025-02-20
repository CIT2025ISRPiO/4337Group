using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Linq;

namespace _4337Project
{
    /// <summary>
    /// Логика взаимодействия для _4337_ShevchukOksana.xaml
    /// </summary>
    public partial class _4337_ShevchukOksana : Window
    {
        string connectionString = @"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=ISRPO2;Integrated Security=True;";

        public _4337_ShevchukOksana()
        {
            InitializeComponent();
        }
        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xls;*.csv";

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                try
                {
                    // контекст лицензии перед использованием EPPlus
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    FileInfo fileInfo = new FileInfo(filePath);
                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Получаем первый лист

                        int rowCount = worksheet.Dimension.Rows;

                        // Строка подключения к SQL Server LocalDB
                        
                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            try
                            {
                                connection.Open();
                                Console.WriteLine("Connection opened!");

                                // Создание таблицы Clients (если она не существует)
                                string createTableQuery = @"
                                    IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Clients')
                                    BEGIN
                                        CREATE TABLE Clients (
                                            Id INT IDENTITY(1,1) PRIMARY KEY,
                                            FullName NVARCHAR(255),
                                            CodeClient NVARCHAR(255),
                                            BirthDate NVARCHAR(255),
                                            [Index] NVARCHAR(255),
                                            City NVARCHAR(255),
                                            Street NVARCHAR(255),
                                            Home INT,
                                            Kvartira INT,
                                            E_mail NVARCHAR(255)
                                        );
                                    END";

                                using (SqlCommand createTableCommand = new SqlCommand(createTableQuery, connection))
                                {
                                    createTableCommand.ExecuteNonQuery();
                                }

                                // Начинаем считывать данные со второй строки (первая - заголовки)
                                for (int row = 2; row <= rowCount; row++)
                                {
                                    // Проверка, что строка не пустая (все ячейки)
                                    if (string.IsNullOrEmpty(worksheet.Cells[row, 1].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 2].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 3].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 4].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 5].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 6].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 7].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 8].Value?.ToString()) &&
                                        string.IsNullOrEmpty(worksheet.Cells[row, 9].Value?.ToString()))
                                    {
                                        // Если все ячейки в строке пустые, прекращаем чтение
                                        break;
                                    }

                                    // FullName (Столбец A, индекс 1)
                                    string fullName = worksheet.Cells[row, 1].Value?.ToString();

                                    // CodeClient (Столбец B, индекс 2)
                                    string codeClient = worksheet.Cells[row, 2].Value?.ToString();

                                    // BirthDate (Столбец C, индекс 3)
                                    DateTime birthDate = DateTime.MinValue; // Значение по умолчанию
                                    string birthDateValue = worksheet.Cells[row, 3].Value?.ToString();

                                    if (worksheet.Cells[row, 3].Value is double oaDate)
                                    {
                                        birthDate = DateTime.FromOADate(oaDate);
                                    }
                                    else if (!string.IsNullOrEmpty(birthDateValue))
                                    {
                                        if (!DateTime.TryParseExact(birthDateValue, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out birthDate))
                                        {
                                            MessageBox.Show($"Не удалось преобразовать значение '{birthDateValue}' в дату (строка: {row}, столбец: 3). Установлено значение по умолчанию.");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show($"Ячейка даты пуста (строка: {row}, столбец: 3). Установлено значение по умолчанию.");
                                    }

                                    // Index (Столбец D, индекс 4)
                                    string index = worksheet.Cells[row, 4].Value?.ToString();

                                    // City (Столбец E, индекс 5)
                                    string city = worksheet.Cells[row, 5].Value?.ToString();

                                    // Street (Столбец F, индекс 6)
                                    string street = worksheet.Cells[row, 6].Value?.ToString();

                                    // Home (Столбец G, индекс 7)
                                    int home = 0;
                                    string homeValue = worksheet.Cells[row, 7].Value?.ToString();
                                    if (!string.IsNullOrEmpty(homeValue))
                                    {
                                        if (!int.TryParse(homeValue, out home))
                                        {
                                            MessageBox.Show($"Не удалось преобразовать значение '{homeValue}' в целое число (строка: {row}, столбец: 7). Установлено значение по умолчанию.");
                                            home = 0;
                                        }
                                    }

                                    // Kvartira (Столбец H, индекс 8)
                                    int kvartira = 0;
                                    string kvartiraValue = worksheet.Cells[row, 8].Value?.ToString();
                                    if (!string.IsNullOrEmpty(kvartiraValue))
                                    {
                                        if (!int.TryParse(kvartiraValue, out kvartira))
                                        {
                                            MessageBox.Show($"Не удалось преобразовать значение '{kvartiraValue}' в целое число (строка: {row}, столбец: 8). Установлено значение по умолчанию.");
                                            kvartira = 0;
                                        }
                                    }

                                    // E-mail (Столбец I, индекс 9)
                                    string email = worksheet.Cells[row, 9].Value?.ToString();

                                    // SQL-запрос для вставки данных
                                    string insertQuery = @"
                                        INSERT INTO Clients (FullName, CodeClient, BirthDate, [Index], City, Street, Home, Kvartira, E_mail)
                                        VALUES (@FullName, @CodeClient, @BirthDate, @Index, @City, @Street, @Home, @Kvartira, @EMail);";

                                    using (SqlCommand insertCommand = new SqlCommand(insertQuery, connection))
                                    {
                                        insertCommand.Parameters.AddWithValue("@FullName", fullName);
                                        insertCommand.Parameters.AddWithValue("@CodeClient", codeClient);
                                        insertCommand.Parameters.AddWithValue("@BirthDate", birthDate);
                                        insertCommand.Parameters.AddWithValue("@Index", index);
                                        insertCommand.Parameters.AddWithValue("@City", city);
                                        insertCommand.Parameters.AddWithValue("@Street", street);
                                        insertCommand.Parameters.AddWithValue("@Home", home);
                                        insertCommand.Parameters.AddWithValue("@Kvartira", kvartira);
                                        insertCommand.Parameters.AddWithValue("@EMail", email);

                                        insertCommand.ExecuteNonQuery();
                                    }
                                }

                                MessageBox.Show("Данные успешно импортированы!");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Ошибка при работе с базой данных: {ex.Message}");
                            }
                            finally
                            {
                                if (connection.State == System.Data.ConnectionState.Open)
                                {
                                    connection.Close();
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте данных: {ex.Message}");
                }
            }
        }


        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            // Строка подключения к SQL Server LocalDB
            
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xlsx";
            saveFileDialog.Title = "Сохранить данные в Excel";
            saveFileDialog.FileName = "ExportedData.xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    try
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        connection.Open();

                        // Запрос для получения данных из базы данных
                        string selectQuery = "SELECT CodeClient, FullName, E_mail, Street FROM Clients"; // Измените запрос для выбора необходимых столбцов
                        SqlCommand selectCommand = new SqlCommand(selectQuery, connection);

                        using (SqlDataReader reader = selectCommand.ExecuteReader())
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            // Словарь для хранения данных по группам
                            var groupedData = new Dictionary<string, List<string[]>>();

                            // Группировка данных по критерию (по улице)
                            while (reader.Read())
                            {
                                string street = reader["Street"].ToString(); // Группировка по столбцу "Street"
                                var rowData = new string[3]; // Массив для хранения нужных данных

                                rowData[0] = reader["CodeClient"].ToString(); // Код клиента
                                rowData[1] = reader["FullName"].ToString();   // ФИО
                                rowData[2] = reader["E_mail"].ToString();      // E-mail

                                if (!groupedData.ContainsKey(street))
                                {
                                    groupedData[street] = new List<string[]>();
                                }
                                groupedData[street].Add(rowData);
                            }

                            // Запись данных на отдельные листы
                            foreach (var group in groupedData)
                            {
                                string groupName = group.Key;
                                var worksheet = package.Workbook.Worksheets.Add(groupName);

                                // Запись заголовков столбцов
                                worksheet.Cells[1, 1].Value = "Код клиента";
                                worksheet.Cells[1, 2].Value = "ФИО";
                                worksheet.Cells[1, 3].Value = "E-mail";

                                // Сортировка данных по ФИО в алфавитном порядке
                                var sortedRows = group.Value.OrderBy(row => row[1]).ToList();

                                // Запись отсортированных данных группы
                                int rowNumber = 2;
                                foreach (var dataRow in sortedRows)
                                {
                                    for (int i = 0; i < dataRow.Length; i++)
                                    {
                                        worksheet.Cells[rowNumber, i + 1].Value = dataRow[i];
                                    }
                                    rowNumber++;
                                }
                            }

                            // Сохранение файла Excel
                            FileInfo excelFile = new FileInfo(saveFileDialog.FileName);
                            package.SaveAs(excelFile);

                            MessageBox.Show("Данные успешно экспортированы в Excel!");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при экспорте данных: {ex.Message}");
                    }
                }
            }
        }
    }
}

