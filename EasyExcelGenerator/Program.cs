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
            foreach (var product in dataStore.Products)
            {
                int row = GetRowForMonth(product.Month);
                int col = GetColumnForProduct(product.Name);

                if (row > 0 && col > 0)
                {
                    worksheet.Cells[row, col].Value = product.SalesAmount;
                }
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
                for (int i = 0; i < 4; i++) // Assuming 4 products
                {
                    barChart.Series[i].XSeries = "A2:A13"; // Months (January - December)
                    barChart.Series[i].Series = $"B2:B13".Replace('B', (char)('B' + i)); // Series for Product A to D
                }
            }
            else
            {
                Console.WriteLine("Bar chart not found.");
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
                pieChart.Series[0].XSeries = "B1:E1";  // Product names (header row)
                pieChart.Series[0].Series = "B13:E13"; // December data for products A to D
            }
            else
            {
                Console.WriteLine("Pie chart not found.");
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
                // Define the X-axis (Months) and Y-axis (Sales Data for Products)
                lineChart.Series[0].XSeries = "A2:A13";  // Months (January - December)
                lineChart.Series[0].Series = "B2:B13";   // Sales data for Product A

                lineChart.Series[1].XSeries = "A2:A13";  // Months (January - December)
                lineChart.Series[1].Series = "C2:C13";   // Sales data for Product B

                lineChart.Series[2].XSeries = "A2:A13";  // Months (January - December)
                lineChart.Series[2].Series = "D2:D13";   // Sales data for Product C

                lineChart.Series[3].XSeries = "A2:A13";  // Months (January - December)
                lineChart.Series[3].Series = "E2:E13";   // Sales data for Product D
            }
            else
            {
                Console.WriteLine("Line chart not found.");
            }
        }
    }

}