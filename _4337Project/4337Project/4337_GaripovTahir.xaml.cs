using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;
using System.IO;
using System.Globalization;

namespace _4337Project
{
    public partial class _4337_GaripovTahir : Window
    {
        private List<RentalRecord> rentalRecords = new List<RentalRecord>();

        public _4337_GaripovTahir()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                rentalRecords.Clear();

                try
                {
                    FileInfo fileInfo = new FileInfo(filePath);
                    using (var package = new ExcelPackage(fileInfo))
                    {
                        if (package.Workbook.Worksheets.Count == 0)
                        {
                            MessageBox.Show("Файл не содержит листов.");
                            return;
                        }

                        var worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension?.Rows ?? 0;
                        if (rowCount == 0)
                        {
                            MessageBox.Show("Файл пустой.");
                            return;
                        }

                        for (int row = 2; row <= rowCount; row++)
                        {
                            try
                            {
                                var record = new RentalRecord
                                {
                                    Id = Convert.ToInt32(worksheet.Cells[row, 1].Text),
                                    OrderCode = worksheet.Cells[row, 2].Text.Trim(),
                                    CreationDate = ParseDate(worksheet.Cells[row, 3].Text),
                                    OrderTime = worksheet.Cells[row, 4].Text.Trim(),
                                    ClientCode = worksheet.Cells[row, 5].Text.Trim(),
                                    Service = worksheet.Cells[row, 6].Text.Trim(),
                                    Status = worksheet.Cells[row, 7].Text.Trim(),
                                    CloseDate = ParseNullableDate(worksheet.Cells[row, 8].Text),
                                    RentalTime = ParseRentalTime(worksheet.Cells[row, 9].Text)
                                };

                                rentalRecords.Add(record);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Ошибка в строке {row}: {ex.Message}");
                            }
                        }
                    }

                    MessageBox.Show($"Успешно импортировано {rentalRecords.Count} записей.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при импорте: {ex.Message}");
                }
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            if (rentalRecords.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта.");
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                try
                {
                    FileInfo newFile = new FileInfo(filePath);
                    if (newFile.Exists) newFile.Delete();

                    using (var package = new ExcelPackage(newFile))
                    {
                        var groupedData = rentalRecords.GroupBy(r => r.CreationDate.ToString("yyyy-MM-dd"));

                        foreach (var group in groupedData)
                        {
                            var worksheet = package.Workbook.Worksheets.Add(group.Key);

                            string[] headers = { "ID", "Код заказа", "Дата создания", "Код клиента", "Услуги" };
                            for (int i = 0; i < headers.Length; i++)
                            {
                                worksheet.Cells[1, i + 1].Value = headers[i];
                            }

                            int row = 2;
                            foreach (var record in group)
                            {
                                worksheet.Cells[row, 1].Value = record.Id;
                                worksheet.Cells[row, 2].Value = "'" + record.OrderCode;
                                worksheet.Cells[row, 3].Value = record.CreationDate.ToShortDateString();
                                worksheet.Cells[row, 4].Value = record.ClientCode;
                                worksheet.Cells[row, 5].Value = record.Service;
                                row++;
                            }
                        }

                        package.Save();
                    }

                    MessageBox.Show($"Успешно экспортировано {rentalRecords.Count} записей.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте: {ex.Message}");
                }
            }
        }


        // Парсинг даты
        private DateTime ParseDate(string value)
        {
            if (DateTime.TryParseExact(value, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
                return result;

            if (double.TryParse(value, out double oaDate))
                return DateTime.FromOADate(oaDate);

            throw new Exception($"Неверный формат даты: {value}");
        }

        private DateTime? ParseNullableDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return ParseDate(value);
        }

        private TimeSpan ParseRentalTime(string rentalTime)
        {
            if (string.IsNullOrWhiteSpace(rentalTime)) return TimeSpan.Zero;

            rentalTime = rentalTime.ToLower().Trim();

            if (rentalTime.Contains("мин"))
            {
                rentalTime = new string(rentalTime.Where(char.IsDigit).ToArray());
                if (int.TryParse(rentalTime, out int minutes))
                    return TimeSpan.FromMinutes(minutes);
            }
            else if (rentalTime.Contains("час"))
            {
                rentalTime = new string(rentalTime.Where(char.IsDigit).ToArray());
                if (int.TryParse(rentalTime, out int hours))
                    return TimeSpan.FromHours(hours);
            }

            return TimeSpan.Zero;
        }

        private string FormatRentalTime(TimeSpan rentalTime)
        {
            if (rentalTime.TotalHours >= 1)
                return $"{(int)rentalTime.TotalHours} часов";
            return $"{(int)rentalTime.TotalMinutes} минут";
        }
    }

    public class RentalRecord
    {
        public int Id { get; set; }
        public string OrderCode { get; set; }
        public DateTime CreationDate { get; set; }
        public string OrderTime { get; set; }
        public string ClientCode { get; set; }
        public string Service { get; set; }
        public string Status { get; set; }
        public DateTime? CloseDate { get; set; }
        public TimeSpan RentalTime { get; set; }
    }
}
