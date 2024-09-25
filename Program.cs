using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelJsonApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // File paths
            string baseDirectory = "/Users/loknathbaskar/Projects/IO_Files/ExcelJson";
            string jsonInputPath = Path.Combine(baseDirectory,"input.json"); 
            string excelOutputPath = Path.Combine(baseDirectory,"output.xlsx");
            string jsonOutputPath = Path.Combine(baseDirectory,"output.json");

            // Step 1: Read JSON and convert to Excel
            ConvertJsonToExcel(jsonInputPath, excelOutputPath);

            // Step 2: Read Excel and convert back to JSON
            ConvertExcelToJson(excelOutputPath, jsonOutputPath);

            Console.WriteLine("JSON to Excel and back to JSON completed successfully.");
        }

        // Converts JSON to Excel file using ClosedXML
        public static void ConvertJsonToExcel(string jsonFilePath, string excelFilePath)
        {
            // Read JSON data from the local file
            string jsonData = File.ReadAllText(jsonFilePath);
            var records = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);

            // Create Excel file using ClosedXML
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Write header row and apply styling (background color)
                var headers = records[0].Keys; // Get the keys from the first record as headers
                int col = 1;
                foreach (var header in headers)
                {
                    var cell = worksheet.Cell(1, col);
                    cell.Value = header;

                    // Apply background color (light blue) and bold formatting to the header
                    cell.Style.Fill.BackgroundColor = XLColor.LightBlue; // Set background color to light blue
                    cell.Style.Font.Bold = true;                         // Set font to bold

                    col++;
                }

                // Write data rows and highlight specific columns
                int row = 2;
                foreach (var record in records)
                {
                    col = 1;
                    foreach (var value in record.Values)
                    {
                        var cell = worksheet.Cell(row, col);
                        cell.Value = value?.ToString() ?? ""; // Handles null values
                        
                        // Highlight columns 2 and 3 with yellow background
                        if (col == 2 || col == 3)
                        {
                            cell.Style.Fill.BackgroundColor = XLColor.Yellow;
                        }

                        col++;
                    }
                    row++;
                }

                // Set auto-fit for all columns
                worksheet.Columns().AdjustToContents(); // Adjust column width to fit content

                // Save the Excel file to the local machine
                workbook.SaveAs(excelFilePath);
            }

            Console.WriteLine($"Excel file generated at: {excelFilePath}");
        }

        // Converts Excel back to JSON file using ClosedXML
        public static void ConvertExcelToJson(string excelFilePath, string jsonFilePath)
        {
            var records = new List<Dictionary<string, object>>();

            // Read Excel file using ClosedXML
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(1); // Assumes the data is in the first worksheet
                var rowCount = worksheet.LastRowUsed().RowNumber();
                var colCount = worksheet.LastColumnUsed().ColumnNumber();

                // Read headers (from the first row)
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cell(1, col).GetString());
                }

                // Read data rows (starting from the second row)
                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new Dictionary<string, object>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        record[headers[col - 1]] = worksheet.Cell(row, col).Value.ToString() ?? ""; // Converts XLCellValue to string and handles nulls
                    }
                    records.Add(record);
                }
            }

            // Serialize data to JSON format and write to the local file
            string jsonData = JsonConvert.SerializeObject(records, Formatting.Indented);
            File.WriteAllText(jsonFilePath, jsonData);

            Console.WriteLine($"JSON file generated at: {jsonFilePath}");
        }
    }
}
