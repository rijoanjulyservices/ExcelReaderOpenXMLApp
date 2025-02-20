using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelReaderApp
{
    public class ReadExcelWithOpenXml
    {
        public   DataTable ReadExcelOpenXml(string filePath)
        {
            var dataTable = new DataTable();

            // Open the Excel file
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();

                if (sheet == null)
                {
                    throw new Exception("No sheets found in the Excel file.");
                }

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Add columns to DataTable based on the first two header rows
                var headerRows = sheetData.Elements<Row>().Take(8).ToList();
                if (headerRows.Count < 8)
                {
                    throw new Exception("Not enough header rows found in the Excel sheet.");
                }

                // Determine the maximum number of columns in the header rows
                int maxColumnIndex = headerRows.SelectMany(row => row.Elements<Cell>())
                                               .Max(cell => GetColumnIndex(cell.CellReference));

                for (int colIndex = 0; colIndex <= maxColumnIndex; colIndex++)
                {
                    string firstHeader = GetCellValue(document, headerRows[0].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == colIndex));
                    string secondHeader = GetCellValue(document, headerRows[1].Elements<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference) == colIndex));

                    // Combine the headers, using a space if either is null
                    string combinedHeader = $"{firstHeader ?? string.Empty} {secondHeader ?? string.Empty}".Trim();
                    if (string.IsNullOrEmpty(combinedHeader))
                    {
                        combinedHeader = $"Column{dataTable.Columns.Count + 1}"; // Default column name if combined header is empty
                    }
                    dataTable.Columns.Add(combinedHeader);
                }

                // Get merged cells information
                var mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

                // Add rows to DataTable, skipping the first two header rows
                foreach (Row row in sheetData.Elements<Row>().Skip(2))
                {
                    var dataRow = dataTable.NewRow();
                    int columnIndex = 0;

                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        // Get the column index of the cell
                        int cellColumnIndex = GetColumnIndex(cell.CellReference);

                        // Ensure the column index is within the bounds of the DataTable
                        if (cellColumnIndex >= dataTable.Columns.Count)
                        {
                            continue;
                        }

                        // Fill in any missing columns with empty values
                        while (columnIndex < cellColumnIndex)
                        {
                            dataRow[columnIndex] = string.Empty;
                            columnIndex++;
                        }

                        // Get the cell value
                        string cellValue = GetCellValue(document, cell);
                        dataRow[columnIndex] = cellValue ?? string.Empty;
                        columnIndex++;
                    }

                    // Ensure the DataRow has the same number of columns as the DataTable
                    while (columnIndex < dataTable.Columns.Count)
                    {
                        dataRow[columnIndex] = string.Empty;
                        columnIndex++;
                    }

                    dataTable.Rows.Add(dataRow);
                }

                // Handle merged cells
                if (mergeCells != null)
                {
                    foreach (MergeCell mergeCell in mergeCells.Elements<MergeCell>())
                    {
                        string[] cellRange = mergeCell.Reference.Value.Split(':');
                        string startCell = cellRange[0];
                        string endCell = cellRange[1];

                        int startColumnIndex = GetColumnIndex(startCell);
                        int endColumnIndex = GetColumnIndex(endCell);
                        int startRowIndex = GetRowIndex(startCell);
                        int endRowIndex = GetRowIndex(endCell);

                        string mergedValue = dataTable.Rows[startRowIndex - 1][startColumnIndex].ToString();

                        for (int rowIndex = startRowIndex - 1; rowIndex < endRowIndex; rowIndex++)
                        {
                            for (int colIndex = startColumnIndex; colIndex <= endColumnIndex; colIndex++)
                            {
                                dataTable.Rows[rowIndex][colIndex] = mergedValue;
                            }
                        }
                    }
                }
            }
            // Remove the first 7 rows
            for (int i = 0; i < 6; i++)
            {
                if (dataTable.Rows.Count > 0)
                {
                    dataTable.Rows[0].Delete();
                }
                dataTable.AcceptChanges();
            }

            

            // Merge the values of the first two rows to create new column names
            var columnNames = new Dictionary<string, int>();
            string previousColumnName = string.Empty;

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                string value1 = dataTable.Rows[0][i] != DBNull.Value ? dataTable.Rows[0][i].ToString().Trim() : "";
                string value2 = dataTable.Rows[1][i] != DBNull.Value ? dataTable.Rows[1][i].ToString().Trim() : $"col{i}";

                // Replace specific values
                value2 = value2.Replace("(JULY Records - Do Not Change)", "July");

                // Use regular expression to replace unwanted characters
                string newColumnName = $"{value1}{value2}";
                newColumnName = Regex.Replace(newColumnName, @"[ \-\/\?']", "");

                // Handle the special case for "Enter Your Records Here (if different)"
                if (value2.Contains("Enter Your Records Here (if different)"))
                {
                    string previousValue1 = dataTable.Rows[0][i - 1] != DBNull.Value ? dataTable.Rows[0][i - 1].ToString().Trim() : "";
                    previousValue1 = Regex.Replace(previousValue1, @"[ \-\/\?']", "");
                    newColumnName = $"{previousValue1}_Client";
                    value2 = newColumnName;
                }
                else if (value2.Contains("Enter Your Records Here"))
                {
                    string previousValue1 = dataTable.Rows[0][i - 1] != DBNull.Value ? $"{value1}" : "";
                    previousValue1 = Regex.Replace(previousValue1, @"[ \-\/\?']", "");
                    newColumnName = $"{previousValue1}_Client";
                }

                // Check for duplicate column names and append a unique identifier if necessary
                if (columnNames.ContainsKey(newColumnName))
                {
                    columnNames[newColumnName]++;
                    newColumnName = $"{newColumnName}_{columnNames[newColumnName]}";
                }
                else
                {
                    columnNames[newColumnName] = 1;
                }

                dataTable.Columns[i].ColumnName = newColumnName;
                previousColumnName = newColumnName; // Update the previous column name
            }

            // Remove the first two rows as they are now used as column names
            dataTable.Rows[0].Delete();
            dataTable.Rows[1].Delete();
            dataTable.AcceptChanges();

            ////// Merge the values of the first two rows to create new column names
            ////for (int i = 0; i < dataTable.Columns.Count; i++)
            ////{
            ////    string value1 = dataTable.Rows[0][i] != DBNull.Value ? dataTable.Rows[0][i].ToString().Trim() : "";
            ////    string value2 = dataTable.Rows[1][i] != DBNull.Value ? dataTable.Rows[1][i].ToString().Trim() : $"col{i}";
            ////    string newColumnName = $"{value1}{value2}".Replace(" ", "").Replace("-", "_").Replace("/", "_").Replace("EnterYourRecordsHere", "_User");
            ////    dataTable.Columns[i].ColumnName = newColumnName;
            ////}

            ////// Remove the first two rows as they are now used as column names
            ////dataTable.Rows[0].Delete();
            ////dataTable.Rows[1].Delete();
            ////dataTable.AcceptChanges();

            return dataTable;
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
