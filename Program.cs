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
                string filePath = @"C:\Users\rijoan\Downloads\JFK Electric, LLC 401(k) Plan_CensusDataImport_09.30.2024.xlsm";
                //string filePath = @"C:\Users\rijoan\Downloads\Cangelose Financial, LLC 401(k) Plan_CensusDataImport_12.31.2024.xlsm";

                // Read the Excel file into a DataTable
                var excelReader = new ReadExcelWithOpenXml();
                var dataTable = excelReader.ReadExcelOpenXml(filePath);

                // Generate table name (e.g., based on file name)
                var tableName = Path.GetFileNameWithoutExtension(filePath);
                tableName = Regex.Replace(tableName, @"[ .(),]", "");

                // Initialize services
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

        ////public static DataTable ReadExcelWithOpenXml(string filePath)
        ////{
        ////    var dataTable = new DataTable();

        ////    // Open the Excel file
        ////    using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
        ////    {
        ////        WorkbookPart workbookPart = document.WorkbookPart;
        ////        Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();

        ////        if (sheet == null)
        ////        {
        ////            throw new Exception("No sheets found in the Excel file.");
        ////        }

        ////        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        ////        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        ////        // Add columns to DataTable based on the first two header rows
        ////        var headerRows = sheetData.Elements<Row>().Take(2).ToList();
        ////        if (headerRows.Count < 2)
        ////        {
        ////            throw new Exception("Not enough header rows found in the Excel sheet.");
        ////        }

        ////        // Determine the maximum number of columns in the header rows
        ////        int maxColumnIndex = headerRows.SelectMany(row => row.Elements<Cell>())
        ////                                       .Max(cell => GetColumnIndex(cell.CellReference));

        ////        for (int colIndex = 0; colIndex <= maxColumnIndex; colIndex++)
        ////        {
        ////            string firstHeader = GetCellValue(document, headerRows[0].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == colIndex));
        ////            string secondHeader = GetCellValue(document, headerRows[1].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == colIndex));

        ////            // Combine the headers, using a space if either is null
        ////            string combinedHeader = $"{firstHeader ?? string.Empty} {secondHeader ?? string.Empty}".Trim();
        ////            if (string.IsNullOrEmpty(combinedHeader))
        ////            {
        ////                combinedHeader = $"Column{dataTable.Columns.Count + 1}"; // Default column name if combined header is empty
        ////            }
        ////            dataTable.Columns.Add(combinedHeader);
        ////        }

        ////        // Get merged cells information
        ////        var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

        ////        // Add rows to DataTable, skipping the first two header rows
        ////        foreach (Row row in sheetData.Elements<Row>().Skip(2))
        ////        {
        ////            var dataRow = dataTable.NewRow();
        ////            int columnIndex = 0;

        ////            foreach (Cell cell in row.Elements<Cell>())
        ////            {
        ////                // Get the column index of the cell
        ////                int cellColumnIndex = GetColumnIndex(cell.CellReference);

        ////                // Ensure the column index is within the bounds of the DataTable
        ////                if (cellColumnIndex >= dataTable.Columns.Count)
        ////                {
        ////                    continue;
        ////                }

        ////                // Fill in any missing columns with empty values
        ////                while (columnIndex < cellColumnIndex)
        ////                {
        ////                    dataRow[columnIndex] = string.Empty;
        ////                    columnIndex++;
        ////                }

        ////                // Get the cell value
        ////                string cellValue = GetCellValue(document, cell);
        ////                dataRow[columnIndex] = cellValue ?? string.Empty;
        ////                columnIndex++;
        ////            }

        ////            // Ensure the DataRow has the same number of columns as the DataTable
        ////            while (columnIndex < dataTable.Columns.Count)
        ////            {
        ////                dataRow[columnIndex] = string.Empty;
        ////                columnIndex++;
        ////            }

        ////            dataTable.Rows.Add(dataRow);
        ////        }

        ////        // Handle merged cells
        ////        if (mergeCells != null)
        ////        {
        ////            foreach (MergeCell mergeCell in mergeCells.Elements<MergeCell>())
        ////            {
        ////                string[] cellRange = mergeCell.Reference.Value.Split(':');
        ////                string startCell = cellRange[0];
        ////                string endCell = cellRange[1];

        ////                int startColumnIndex = GetColumnIndex(startCell);
        ////                int endColumnIndex = GetColumnIndex(endCell);
        ////                int startRowIndex = GetRowIndex(startCell);
        ////                int endRowIndex = GetRowIndex(endCell);

        ////                string mergedValue = dataTable.Rows[startRowIndex - 1][startColumnIndex].ToString();

        ////                for (int rowIndex = startRowIndex - 1; rowIndex < endRowIndex; rowIndex++)
        ////                {
        ////                    for (int colIndex = startColumnIndex; colIndex <= endColumnIndex; colIndex++)
        ////                    {
        ////                        dataTable.Rows[rowIndex][colIndex] = mergedValue;
        ////                    }
        ////                }
        ////            }
        ////        }
        ////    }

        ////    return dataTable;
        ////}

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