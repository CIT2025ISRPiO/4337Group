
    using System.Data;
using System.IO;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Linq;

namespace _4337Project


    {


    public class Service
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public string ServiceType { get; set; }
        public string ServiceCode { get; set; }
        public decimal Price { get; set; }
    }


    public partial class _4337_Baryshev : Window
        {
            string connectionString = "Server=DESKTOP-FVJO5QP; Database=serviceForExcel; Integrated Security=True; TrustServerCertificate=True;";

            public _4337_Baryshev()
            {
                InitializeComponent();
                CreateDatabaseTable();
            }

            private void CreateDatabaseTable()
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string createTableQuery = @"
                    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='Services' AND xtype='U')
                    CREATE TABLE Services (
                        Id INT PRIMARY KEY,
                        Title NVARCHAR(255),
                        ServiceType NVARCHAR(255),
                        ServiceCode NVARCHAR(255),
                        Price DECIMAL(18,2)
                    )";
                    new SqlCommand(createTableQuery, connection).ExecuteNonQuery();
                }
            }

            // Импорт данных
            private void ImportButton_Click(object sender, RoutedEventArgs e)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xlsx";
                if (openFileDialog.ShowDialog() == true)
                {
                    using (var workbook = new XLWorkbook(openFileDialog.FileName))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var range = worksheet.RangeUsed();

                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            SqlCommand truncateCmd = new SqlCommand("DELETE FROM Services", connection);
                            truncateCmd.ExecuteNonQuery();

                            foreach (var row in range.Rows().Skip(1))
                            {
                                string insertQuery = @"
                                INSERT INTO Services 
                                VALUES (@Id, @Title, @ServiceType, @ServiceCode, @Price)";

                                SqlCommand cmd = new SqlCommand(insertQuery, connection);
                                cmd.Parameters.AddWithValue("@Id", row.Cell(1).GetValue<int>());
                                cmd.Parameters.AddWithValue("@Title", row.Cell(2).GetValue<string>());
                                cmd.Parameters.AddWithValue("@ServiceType", row.Cell(3).GetValue<string>());
                                cmd.Parameters.AddWithValue("@ServiceCode", row.Cell(4).GetValue<string>());
                                cmd.Parameters.AddWithValue("@Price", row.Cell(5).GetValue<decimal>());
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    MessageBox.Show("Данные успешно импортированы!");
                }
            }

            // Экспорт данных
            private void ExportButton_Click(object sender, RoutedEventArgs e)
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    string query = "SELECT * FROM Services ORDER BY ServiceType, Price";
                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    using (var workbook = new XLWorkbook())
                    {
                        foreach (var group in dt.AsEnumerable().GroupBy(r => r["ServiceType"]))
                        {
                            var ws = workbook.Worksheets.Add(group.Key.ToString());
                            ws.Cell(1, 1).Value = "ID";
                            ws.Cell(1, 2).Value = "Название услуги";
                            ws.Cell(1, 3).Value = "Стоимость";

                            int row = 2;
                            foreach (var item in group.OrderBy(x => x["Price"]))
                            {
                                ws.Cell(row, 1).Value = item["Id"].ToString();
                                ws.Cell(row, 2).Value = item["Title"].ToString();
                                ws.Cell(row, 3).Value = item["Price"].ToString();
                                row++;
                            }
                        }

                        SaveFileDialog saveDialog = new SaveFileDialog();
                        saveDialog.Filter = "Excel Files|*.xlsx";
                        if (saveDialog.ShowDialog() == true)
                        {
                            workbook.SaveAs(saveDialog.FileName);
                            MessageBox.Show("Экспорт завершён!");
                        }
                    }
                }
            }
        }
    }


    /// <summary>
    /// Логика взаимодействия для _4337_Baryshev.xaml
    /// </summary>
