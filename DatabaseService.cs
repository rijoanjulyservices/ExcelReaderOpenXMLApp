using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderApp
{
    public class DatabaseService
    {
        private readonly string _connectionString;

        public DatabaseService(string connectionString)
        {
            _connectionString = connectionString;
        }

        public void CreateTable(string createTableScript)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(createTableScript, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        public void BulkInsertData(string tableName, DataTable dataTable)
        {
            // Remove rows where all columns are blank or null
            RemoveEmptyRows(dataTable);

            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (var bulkCopy = new SqlBulkCopy(connection))
                {
                    bulkCopy.DestinationTableName = tableName;

                    // Map columns by name (Excel -> SQL)
                    foreach (DataColumn column in dataTable.Columns)
                    {
                        bulkCopy.ColumnMappings.Add(column.ColumnName, column.ColumnName);
                    }

                    bulkCopy.WriteToServer(dataTable);
                }
            }
        }

        private void RemoveEmptyRows(DataTable dataTable)
        {
            for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
            {
                var row = dataTable.Rows[i];
                bool isEmpty = true;

                foreach (var item in row.ItemArray)
                {
                    if (item != null && !string.IsNullOrWhiteSpace(item.ToString()))
                    {
                        isEmpty = false;
                        break;
                    }
                }

                if (isEmpty)
                {
                    dataTable.Rows.Remove(row);
                }
            }
        }

        public bool TableExists(string tableName)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                var command = new SqlCommand(
                    $"SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = @TableName",
                    connection);
                command.Parameters.AddWithValue("@TableName", tableName);
                return command.ExecuteScalar() != null;
            }
        }
        public void ExecuteCommand(string commandText)
        {
            using (var connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                using (var command = new SqlCommand(commandText, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
