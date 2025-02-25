using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace _4337Project
{
    public class Vafin
    {
        public static void ImportData(string filePath, string connectionString, string tableName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo fileInfo = new FileInfo(filePath);
            if (!fileInfo.Exists) throw new FileNotFoundException("Файл не найден.", filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                if (package.Workbook.Worksheets.Count == 0)
                    throw new InvalidOperationException("Файл не содержит ни одного листа.");

                ExcelWorksheet worksheet = package.Workbook.Worksheets["Лист1"];
                worksheet.Calculate();
                int rowCount = worksheet.Dimension?.Rows ?? 0;

                if (rowCount == 0) throw new InvalidOperationException("Лист пустой.");

                CreateTable(connectionString, tableName);
                SaveDataToTable(connectionString, tableName, worksheet, rowCount);
            }
        }

        private static void CreateTable(string connectionString, string tableName)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string createTableQuery = $@"
                IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = '{tableName}')
                BEGIN
                    CREATE TABLE {tableName} (
                        Id INT IDENTITY(1,1) PRIMARY KEY,
                        [Код заказа] NVARCHAR(50),
                        [Дата создания] DATE,
                        [Время заказа] TIME,
                        [Код клиента] NVARCHAR(50),
                        [Услуги] NVARCHAR(MAX),
                        [Статус] NVARCHAR(50),
                        [Дата закрытия] DATE,
                        [Время проката] NVARCHAR(50)
                    )
                END";
                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private static void SaveDataToTable(string connectionString, string tableName, ExcelWorksheet worksheet, int rowCount)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                for (int row = 2; row <= rowCount; row++)
                {
                    string orderCode = worksheet.Cells[row, 2].Text.Trim(); 
                    string orderDateText = worksheet.Cells[row, 3].Text.Trim(); 
                    string orderTimeText = worksheet.Cells[row, 4].Text.Trim(); 
                    string clientCode = worksheet.Cells[row, 5].Text.Trim(); 
                    string services = worksheet.Cells[row, 6].Text.Trim(); 
                    string status = worksheet.Cells[row, 7].Text.Trim(); 
                    string closeDateText = worksheet.Cells[row, 8].Text.Trim(); 
                    string rentalTime = worksheet.Cells[row, 9].Text.Trim(); 

                   
                    DateTime? orderDate = ParseDate(orderDateText);
                    TimeSpan? orderTime = ParseTime(orderTimeText);
                    DateTime? closeDate = ParseDate(closeDateText);

                   
                    string insertQuery = $@"
                    INSERT INTO {tableName} (
                        [Код заказа], [Дата создания], [Время заказа], [Код клиента], 
                        [Услуги], [Статус], [Дата закрытия], [Время проката]
                    ) VALUES (@p1, @p2, @p3, @p4, @p5, @p6, @p7, @p8)";

                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        command.Parameters.AddWithValue("@p1", orderCode ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p2", orderDate ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p3", orderTime ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p4", clientCode ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p5", services ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p6", status ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p7", closeDate ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@p8", rentalTime ?? (object)DBNull.Value);
                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        public static void ExportData(string connectionString, string tableName, string outputFilePath)
        {
            List<Dictionary<string, object>> data = GetDataFromTable(connectionString, tableName);
            var groupedData = data.GroupBy(row => row["Статус"]?.ToString() ?? "Нет статуса"); 
            CreateExcel(groupedData, outputFilePath);
        }

        private static List<Dictionary<string, object>> GetDataFromTable(string connectionString, string tableName)
        {
            List<Dictionary<string, object>> data = new List<Dictionary<string, object>>();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                string query = $"SELECT * FROM {tableName}";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Dictionary<string, object> row = new Dictionary<string, object>();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                string columnName = reader.GetName(i);
                                row[columnName] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                            }
                            data.Add(row);
                        }
                    }
                }
            }
            return data;
        }

        private static void CreateExcel(IEnumerable<IGrouping<string, Dictionary<string, object>>> groupedData, string outputFilePath)
        {
            FileInfo newFile = new FileInfo(outputFilePath);
            if (newFile.Exists) newFile.Delete();

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                foreach (var group in groupedData)
                {
                    string sheetName = string.IsNullOrEmpty(group.Key) ? "Без статуса" : group.Key;
                    sheetName = sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;

                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                    string[] columnNames = { "Id", "Код заказа", "Дата создания", "Код клиента", "Услуги" };

                    for (int i = 0; i < columnNames.Length; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = columnNames[i];
                    }

                    int row = 2;
                    foreach (var record in group)
                    {
                        worksheet.Cells[row, 1].Value = record["Id"] ?? "Нет данных";
                        worksheet.Cells[row, 2].Value = record["Код заказа"] ?? "Нет данных";

                        var creationDate = record["Дата создания"];
                        if (creationDate != null && DateTime.TryParse(creationDate.ToString(), out DateTime parsedDate))
                        {
                            worksheet.Cells[row, 3].Value = parsedDate;
                            worksheet.Cells[row, 3].Style.Numberformat.Format = "yyyy-mm-dd"; 
                        }
                        else
                        {
                            worksheet.Cells[row, 3].Value = "Нет данных";
                        }

                        worksheet.Cells[row, 4].Value = record["Код клиента"] ?? "Нет данных";
                        worksheet.Cells[row, 5].Value = record["Услуги"] ?? "Нет данных";

                        row++;
                    }

                    
                    worksheet.Column(3).AutoFit(); 
                }
                package.Save();
            }
        }





        private static DateTime? ParseDate(string dateText)
        {
            if (string.IsNullOrEmpty(dateText)) return null;
            DateTime parsedDate;
            if (DateTime.TryParse(dateText, out parsedDate))
            {
                return parsedDate;
            }
            return null;
        }

        private static TimeSpan? ParseTime(string timeText)
        {
            if (string.IsNullOrEmpty(timeText)) return null;
            TimeSpan parsedTime;
            if (TimeSpan.TryParse(timeText, out parsedTime))
            {
                return parsedTime;
            }
            return null;
        }
    }
}
