
using System.Data;
using System.IO;
using System.Windows;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Linq;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System;

namespace _4337Project


    {


    public class Service
    {
        [JsonProperty("IdServices")]
        public int Id { get; set; }

        [JsonProperty("NameServices")]
        public string Title { get; set; }

        [JsonProperty("TypeOfService")]
        public string ServiceType { get; set; }

        [JsonProperty("CodeService")]
        public string ServiceCode { get; set; }

        [JsonProperty("Cost")]
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

            private void ImportJsonButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "JSON Files|*.json";
            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    string json = File.ReadAllText(openFileDialog.FileName);
                    var services = JsonConvert.DeserializeObject<List<Service>>(json);

                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        SqlCommand truncateCmd = new SqlCommand("DELETE FROM Services", connection);
                        truncateCmd.ExecuteNonQuery();

                        foreach (var service in services)
                        {
                            string insertQuery = @"
                    INSERT INTO Services 
                    VALUES (@Id, @Title, @ServiceType, @ServiceCode, @Price)";

                            SqlCommand cmd = new SqlCommand(insertQuery, connection);
                            cmd.Parameters.AddWithValue("@Id", service.Id);
                            cmd.Parameters.AddWithValue("@Title", service.Title);
                            cmd.Parameters.AddWithValue("@ServiceType", service.ServiceType);
                            cmd.Parameters.AddWithValue("@ServiceCode", service.ServiceCode);
                            cmd.Parameters.AddWithValue("@Price", service.Price);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    MessageBox.Show("Данные из JSON успешно импортированы!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}");
                }
            }
        }

            private void ExportToWordButton_Click(object sender, RoutedEventArgs e)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = "SELECT * FROM Services ORDER BY ServiceType, Price";
                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Word Documents|*.docx";
                if (saveDialog.ShowDialog() == true)
                {
                    using (WordprocessingDocument doc = WordprocessingDocument.Create(saveDialog.FileName, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = doc.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        Body body = mainPart.Document.AppendChild(new Body());

                        foreach (var group in dt.AsEnumerable().GroupBy(r => r["ServiceType"]))
                        {
                            // Заголовок категории
                            Paragraph headerPara = new Paragraph();
                            Run headerRun = new Run();
                            headerRun.AppendChild(new Text(group.Key.ToString()));
                            headerRun.RunProperties = new RunProperties(new Bold());
                            headerPara.AppendChild(headerRun);
                            body.AppendChild(headerPara);

                            // Создание таблицы
                            Table table = new Table();
                            TableProperties props = new TableProperties(
                                new TableBorders(
                                    new TopBorder() { Val = BorderValues.Single },
                                    new BottomBorder() { Val = BorderValues.Single },
                                    new LeftBorder() { Val = BorderValues.Single },
                                    new RightBorder() { Val = BorderValues.Single }
                                )
                            );
                            table.AppendChild(props);

                            // Заголовки таблицы
                            TableRow headerRow = new TableRow();
                            headerRow.Append(
                                CreateCell("ID", true),
                                CreateCell("Название услуги", true),
                                CreateCell("Стоимость", true)
                            );
                            table.AppendChild(headerRow);

                            // Данные
                            foreach (var item in group.OrderBy(x => x["Price"]))
                            {
                                TableRow dataRow = new TableRow();
                                dataRow.Append(
                                    CreateCell(item["Id"].ToString()),
                                    CreateCell(item["Title"].ToString()),
                                    CreateCell(item["Price"].ToString())
                                );
                                table.AppendChild(dataRow);
                            }

                            body.AppendChild(table);
                            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                        }
                    }
                    MessageBox.Show("Экспорт в Word завершён!");
                }
            }
        }

            private TableCell CreateCell(string text, bool isHeader = false)
        {
            TableCell cell = new TableCell();
            Paragraph para = new Paragraph();
            Run run = new Run();
            run.AppendChild(new Text(text));

            if (isHeader)
            {
                run.RunProperties = new RunProperties(new Bold());
            }

            para.AppendChild(run);
            cell.AppendChild(para);
            return cell;
        }


    }
    }


    /// <summary>
    /// Логика взаимодействия для _4337_Baryshev.xaml
    /// </summary>
