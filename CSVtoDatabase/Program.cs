using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using ClosedXML.Excel;

class Program
{

    static void Main(string[] args)
    {

        string excelFilePath = @"C:\Users\nguye\source\repos\CSVtoDatabase\CSVtoDatabase\bin\Debug\net7.0\Orders-With Nulls.xlsx";
        string connectionString = @"Server=localhost\SQLEXPRESS; Database=OrdersWithNulls; Trusted_Connection=True;";
        DataTable table = LoadExcelData(excelFilePath);
        InsertIntoSqlServer(table, connectionString);
        Console.WriteLine("Import complete.");

        static DataTable LoadExcelData(string filePath)
        {
            var dt = new DataTable();

            using (var workbook = new XLWorkbook(filePath))
            {
                var ws = workbook.Worksheet(1);
                var firstRow = true;

                foreach (var row in ws.RowsUsed())
                {
                    if (firstRow)
                    {
                        // Set up DataTable columns based on header row
                        foreach (var cell in row.Cells())
                            dt.Columns.Add(cell.Value.ToString());
                        firstRow = false;
                    }
                    else
                    {
                        // Add a new row and assign cell values
                        DataRow newRow = dt.NewRow();
                        int columnCount = dt.Columns.Count;

                        for (int i = 0; i < columnCount; i++)
                        {
                            var cellValue = row.Cell(i + 1).Value;
                            newRow[i] = string.IsNullOrWhiteSpace(cellValue.ToString()) ? DBNull.Value : (object)cellValue;
                        }

                        dt.Rows.Add(newRow);
                    }
                }

            }
            return dt;
        }

        static void InsertIntoSqlServer(DataTable dt, string connectionString)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                foreach (DataRow row in dt.Rows)
                {
                    string query = @"insert into Orders([Order], [Order Date], [Order Quantity], [Sales], [Ship Mode], [Profit], [Unit Price], [Customer Name], [Customer Segment], [Product Category])

                     values(@Order, @OrderDate, @OrderQuantity, @Sales, @ShipMode, @Profit, @UnitPrice, @CustomerName, @CustomerSegment, @ProductCategory)";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Order", row["Order ID"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@OrderDate", row["Order Date"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@OrderQuantity", row["Order Quantity"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Sales", row["Sales"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@ShipMode", row["Ship Mode"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@Profit", row["Profit"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@UnitPrice", row["Unit Price"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@CustomerName", row["Customer Name"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@CustomerSegment", row["Customer Segment"] ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@ProductCategory", row["Product Category"] ?? DBNull.Value);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}
