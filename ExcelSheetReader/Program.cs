#pragma warning disable S1075, S1215, CA1416

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelSheetReader.Helpers;
using ExcelSheetReader.Settings;

namespace ExcelSheetReader
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            // Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application ExcelApp = new Excel.Application();
            Excel.Workbook ExcelWorkbook = ExcelApp.Workbooks.Open(Config.FilePath);
            Excel._Worksheet ExcelWorksheet = ExcelWorkbook.Sheets
                .Cast<Excel.Worksheet>()
                .First(s => s.Name == Config.SheetName);
            Excel.Range ExcelRange = ExcelWorksheet.UsedRange;

            // Generate and output the SQL query
            string output = QueryGenerator(ExcelRange);
            Console.WriteLine(output);
            
            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Rule of thumb for releasing com objects:
            //   never use two dots, all COM objects must be referenced and released individually
            //   ex: [somthing].[something].[something] is bad

            // Release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(ExcelRange);
            Marshal.ReleaseComObject(ExcelWorksheet);

            // Close and release
            ExcelWorkbook.Close();
            Marshal.ReleaseComObject(ExcelWorkbook);

            // Quit and release
            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelApp);

        }

        private static string QueryGenerator(Excel.Range xlRange)
        {
            string finalQuery = string.Empty;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                List<string> cellValues = RowDataExtractor(i, colCount, xlRange);

                string queryIndent = "    ";
                string queryStart = "Values (";
                string queryMiddle = string.Join(", ", cellValues);
                string queryEnd = ");";
                string queryTerminator = "\r\n";

                string insertQuery = QueryHelper.Queries[Config.SheetName] + queryTerminator;
                string valueQuery = string.Join(
                    string.Empty, new string[] { queryIndent, queryStart, queryMiddle, queryEnd, queryTerminator }
                );

                finalQuery += (insertQuery + valueQuery);
            }

            return finalQuery;
        }


        private static List<string> RowDataExtractor(int currentRow, int colCount, Excel.Range xlRange)
        {
            List<string> outputList = new();

            for (int j = 2; j <= colCount; j++)
            {
                var colTitle = xlRange.Cells[1, j].Value2;

                if (string.IsNullOrWhiteSpace(colTitle))
                {
                    break;
                }

                string? validVal;
                var cell = xlRange.Cells[currentRow, j];
                var cellVal = cell?.Value2;

                bool cellValIsCode = cell != null
                    && cellVal != null
                    && cellVal!.GetType() == typeof(string)
                    && CodeHelper.Codes.ContainsKey(cellVal);

                validVal = cellValIsCode ? $"{CodeHelper.Codes[cellVal]}" : FormatNonCodeData(cellVal ?? "null");

                outputList.Add(validVal);
            }

            return outputList;
        }

        private static string FormatNonCodeData<T>(T cellVal)
        {
            string output = (!cellVal!.Equals("null")) ? $"'{cellVal}'" : "null";
            return output;
        }
    }
}
