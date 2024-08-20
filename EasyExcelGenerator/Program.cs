using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace EasyExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            string basePath = AppDomain.CurrentDomain.BaseDirectory;

            string templatePath = Path.Combine(basePath, @"..\..\..\docs\Invoice_Template.xlsx");
            string newFilePath = Path.Combine(basePath, "Populated_Invoice_With_Charts.xlsx");

            if (!File.Exists(templatePath))
            {
                Console.WriteLine("Template not found!");
                return;
            }

            try
            {
                using (var package = new ExcelPackage(new FileInfo(templatePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    var dataStore = new ProductDataStore();
                    PopulateWorksheetWithData(worksheet, dataStore);

                    // Update the charts
                    var chartUpdater = new ChartUpdater();
                    chartUpdater.UpdateCharts(worksheet);

                    // Save the updated file
                    package.SaveAs(new FileInfo(newFilePath));
                    Console.WriteLine($"Excel file with charts saved to: {newFilePath}");
                }
            }
            catch (Exception ex)
            {
                // In a larger application, replace this with logging to a file or logging service.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <summary>
        /// Populates the worksheet with product sales data.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet to populate.</param>
        /// <param name="dataStore">The source of product sales data.</param>
        private static void PopulateWorksheetWithData(ExcelWorksheet worksheet, ProductDataStore dataStore)
        {
            // Set the header spanning B1 to G1
            worksheet.Cells["B1:G1"].Merge = true;
            worksheet.Cells["B1"].Value = "PRODUCT SALES OVERVIEW";

            // Set column headers in row 2 (B2:G2)
            worksheet.Cells["B2"].Value = "Month";
            worksheet.Cells["C2"].Value = "Product A Sales";
            worksheet.Cells["D2"].Value = "Product B Sales";
            worksheet.Cells["E2"].Value = "Product C Sales";
            worksheet.Cells["F2"].Value = "Product D Sales";
            worksheet.Cells["G2"].Value = "Total Sales";

            // Populate rows with data (starting from row 3)
            int startRow = 3;
            foreach (var month in dataStore.Products.GroupBy(p => p.Month))
            {
                var row = startRow++;
                worksheet.Cells[row, 2].Value = month.Key;  // Month name in column B

                // Add sales data for each product
                worksheet.Cells[row, 3].Value = month.First(p => p.Name == "Product A").SalesAmount;
                worksheet.Cells[row, 4].Value = month.First(p => p.Name == "Product B").SalesAmount;
                worksheet.Cells[row, 5].Value = month.First(p => p.Name == "Product C").SalesAmount;
                worksheet.Cells[row, 6].Value = month.First(p => p.Name == "Product D").SalesAmount;

                // Calculate total sales and place it in column G
                worksheet.Cells[row, 7].Formula = $"SUM(C{row}:F{row})";
            }
        }


        /// <summary>
        /// Gets the row number for the specified month.
        /// </summary>
        /// <param name="month">The month.</param>
        /// <returns>The row number.</returns>
        private static int GetRowForMonth(string month)
        {
            var months = new[] { "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" };
            int rowIndex = Array.IndexOf(months, month);
            return rowIndex >= 0 ? rowIndex + 2 : -1; // +2 because the data starts in row 2
        }

        /// <summary>
        /// Gets the column number for the specified product name.
        /// </summary>
        /// <param name="productName">The product name.</param>
        /// <returns>The column number.</returns>
        private static int GetColumnForProduct(string productName)
        {
            var products = new[] { "Product A", "Product B", "Product C", "Product D" };
            int columnIndex = Array.IndexOf(products, productName);
            return columnIndex >= 0 ? columnIndex + 2 : -1; // +2 because the data starts in column B
        }
    }

    /// <summary>
    /// Handles chart updates in the Excel worksheet.
    /// </summary>
    public class ChartUpdater
    {
        /// <summary>
        /// Updates all charts in the worksheet.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet.</param>
        public void UpdateCharts(ExcelWorksheet worksheet)
        {
            UpdateBarChart(worksheet);
            UpdatePieChart(worksheet);
            UpdateLineChart(worksheet);
        }

        /// <summary>
        /// Updates the bar chart with new data.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet.</param>
        private void UpdateBarChart(ExcelWorksheet worksheet)
        {
            var barChart = worksheet.Drawings["BarChart"] as ExcelBarChart;

            if (barChart != null)
            {
                // Correct references for each product
                barChart.Series[0].XSeries = "B3:B14"; // Months (January - December)
                barChart.Series[0].Series = "C3:C14";  // Product A Sales

                barChart.Series[1].XSeries = "B3:B14"; // Months (January - December)
                barChart.Series[1].Series = "D3:D14";  // Product B Sales

                barChart.Series[2].XSeries = "B3:B14"; // Months (January - December)
                barChart.Series[2].Series = "E3:E14";  // Product C Sales

                barChart.Series[3].XSeries = "B3:B14"; // Months (January - December)
                barChart.Series[3].Series = "F3:F14";  // Product D Sales
            }
        }

        /// <summary>
        /// Updates the pie chart with new data.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet.</param>
        private void UpdatePieChart(ExcelWorksheet worksheet)
        {
            var pieChart = worksheet.Drawings["PieChart"] as ExcelPieChart;

            if (pieChart != null)
            {
                // Correct references for December
                pieChart.Series[0].XSeries = "C2:F2";  // Product headers
                pieChart.Series[0].Series = "C14:F14"; // December data for products A to D
            }
        }

        /// <summary>
        /// Updates the line chart with new data.
        /// </summary>
        /// <param name="worksheet">The Excel worksheet.</param>
        private void UpdateLineChart(ExcelWorksheet worksheet)
        {
            var lineChart = worksheet.Drawings["LineChart"] as ExcelLineChart;

            if (lineChart != null)
            {
                // Correct references for all products
                lineChart.Series[0].XSeries = "B3:B14";  // Months (January - December)
                lineChart.Series[0].Series = "C3:C14";   // Product A Sales

                lineChart.Series[1].XSeries = "B3:B14";  // Months (January - December)
                lineChart.Series[1].Series = "D3:D14";   // Product B Sales

                lineChart.Series[2].XSeries = "B3:B14";  // Months (January - December)
                lineChart.Series[2].Series = "E3:E14";   // Product C Sales

                lineChart.Series[3].XSeries = "B3:B14";  // Months (January - December)
                lineChart.Series[3].Series = "F3:F14";   // Product D Sales
            }
        }
    }
}