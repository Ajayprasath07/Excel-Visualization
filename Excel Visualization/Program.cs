using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Macros
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"C:\Users\Ajayprasath\OneDrive\Desktop\Soustr\Demo Ex.xlsx";

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                Console.WriteLine("Opening Excel file...");
                // Open the Excel file
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets["Sheet1"]; // Adjust sheet name if necessary

                // Define the range for the charts
                Console.WriteLine("Defining range for charts...");
                Excel.Range chartRange = worksheet.Range["B7:C12"];

                // Add a Pie Chart
                Console.WriteLine("Adding Pie Chart...");
                Excel.ChartObject pieChart = worksheet.ChartObjects().Add(100, 50, 300, 300);
                pieChart.Chart.ChartType = Excel.XlChartType.xlPie;
                pieChart.Chart.SetSourceData(chartRange);

                // Add a Clustered Column Chart
                Console.WriteLine("Adding Clustered Column Chart...");
                Excel.ChartObject columnChart = worksheet.ChartObjects().Add(400, 50, 300, 300);
                columnChart.Chart.ChartType = Excel.XlChartType.xlColumnClustered;
                columnChart.Chart.SetSourceData(chartRange);

                // Select a different cell at the end (optional)
                Console.WriteLine("Selecting cell R24...");
                worksheet.Range["R24"].Select();

                // Save the workbook if needed
                Console.WriteLine("Saving workbook...");
                workbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Cleanup
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                workbook = null;
                worksheet = null;
                excelApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                Console.WriteLine("Cleanup complete.");
            }
        }
    }
}

