using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelReaderApp;

namespace ReadExcelOpenXml
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the Excel file
                //string filePath = @"C:\Users\rijoan\Downloads\JFK Electric, LLC 401(k) Plan_CensusDataImport_09.30.2024.xlsm";
                //string filePath = @"C:\Users\rijoan\Downloads\Cangelose Financial, LLC 401(k) Plan_CensusDataImport_12.31.2024.xlsm";
                string filePath = @"C:\Users\rijoan\Desktop\Archive\Files\Admin_Division.xlsm";

                // Read the Excel file into a DataTable
                var excelReader = new ReadExcelWithOpenXml();
                var dataTable = excelReader.ReadExcelOpenXml(filePath);

                // Generate table name (e.g., based on file name)
                var tableName = Path.GetFileNameWithoutExtension(filePath);
                tableName = Regex.Replace(tableName, @"[ .(),] '", "");

                // Initialize services
                //var connectionString = "Server=testsql1.julyservices.local;Database=TPAManager_Test;User ID=report;Password=brlWOyuz07Rljof#e!ug;Integrated Security=true;TrustServerCertificate=True;";
                var connectionString = "Server=SWD-RIJOAN-L;Database=TestDB;User ID=jahangir;Password=Baylor123;Integrated Security=true;TrustServerCertificate=True;";
                var sqlGenerator = new SqlTableGenerator();
                var dbService = new DatabaseService(connectionString);

                // Check if table exists
                if (dbService.TableExists(tableName))
                {
                    Console.WriteLine($"Table '{tableName}' already exists. Do you want to overwrite it? (Y/N)");
                    var response = Console.ReadLine();
                    if (response?.ToUpper() == "Y")
                    {
                        try
                        {
                            // Truncate the table
                            dbService.ExecuteCommand($"TRUNCATE TABLE {tableName}");
                            Console.WriteLine($"Table '{tableName}' truncated.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error executing command: {ex.Message}");
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Operation canceled.");
                        return;
                    }
                }
                else
                {
                    // Generate and execute CREATE TABLE script if the table does not exist
                    var createTableScript = sqlGenerator.GenerateCreateTableScript(dataTable, tableName);
                    dbService.CreateTable(createTableScript);
                }

                // Bulk insert data
                dbService.BulkInsertData(tableName, dataTable);

                Console.WriteLine($"Table '{tableName}' created and data inserted successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }


        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell == null || cell.CellValue == null)
            {
                return null; // Return null to indicate skipping
            }

            SharedStringTablePart sharedStringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (sharedStringTablePart != null && int.TryParse(value, out int index))
                {
                    return sharedStringTablePart.SharedStringTable.ChildElements[index].InnerText;
                }
            }

            return value ?? string.Empty;
        }

        private static int GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return -1; // Return -1 to indicate an invalid column index
            }

            string columnReference = new string(cellReference.Where(char.IsLetter).ToArray());
            int columnIndex = 0;
            int factor = 1;

            for (int i = columnReference.Length - 1; i >= 0; i--)
            {
                columnIndex += factor * (columnReference[i] - 'A' + 1);
                factor *= 26; // Corrected from 66 to 26
            }

            return columnIndex - 1; // Convert to zero-based index
        }

        private static int GetRowIndex(string cellReference)
        {
            string rowReference = new string(cellReference.Where(char.IsDigit).ToArray());
            return int.Parse(rowReference);
        }
    }
}