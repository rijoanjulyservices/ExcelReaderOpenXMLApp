using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace   ExcelReaderApp

{
    public class SqlTableGenerator
    {
        public string GenerateCreateTableScript(DataTable dataTable, string tableName)
        {
            var columns = new List<string>();

            foreach (DataColumn column in dataTable.Columns)
            {
                // Map Excel data types to SQL Server data types
                string sqlType = GetSqlType(column);
                columns.Add($"[{column.ColumnName}] {sqlType}");
            }

            return $"CREATE TABLE {tableName} ({string.Join(", ", columns)});";
        }

        private string GetSqlType(DataColumn column)
        {
            // Check for specific column names first
            //if (column.ColumnName.Equals("TERMDATE", StringComparison.OrdinalIgnoreCase) ||
            //    column.ColumnName.Equals("BIRTHDATE", StringComparison.OrdinalIgnoreCase) ||
            //    column.ColumnName.Equals("HIREDATE", StringComparison.OrdinalIgnoreCase))
            //{
            //    return "DATETIME";
            //}

            // Map .NET types to SQL Server types
            Type dataType = column.DataType;
            if (dataType == typeof(string))
                return "NVARCHAR(255)";
            if (dataType == typeof(int))
                return "INT";
            if (dataType == typeof(DateTime))
                return "DATETIME";
            if (dataType == typeof(decimal) || dataType == typeof(double))
                return "DECIMAL(18, 2)";
            if (dataType == typeof(bool))
                return "BIT";

            // Default to NVARCHAR if type is unknown
            return "NVARCHAR(255)";
        }
    }
}
