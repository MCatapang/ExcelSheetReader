#pragma warning disable S1075, S1215, CA1416

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelSheetReader.Helpers;
using ExcelSheetReader.Settings;
using System.Diagnostics;

namespace ExcelSheetReader
{
    public static class Program
    {
        public static Excel.Application ExcelApp { get; private set; }
        public static Excel.Workbook ExcelWorkbook { get; private set; }
        public static Excel.Sheets ExcelSheets { get; private set; }
        public static IEnumerable<Excel.Worksheet> ExcelWorksheets { get; private set; }
        public static Excel._Worksheet ExcelWorksheet { get; private set; }
        public static Excel.Range ExcelRange { get; private set; }

        static Program()
        {
            // Create COM Objects. Create a COM object for everything that is referenced
            ExcelApp = new Excel.Application();
            ExcelWorkbook = ExcelApp.Workbooks.Open(Config.FilePath);
            ExcelSheets = ExcelWorkbook.Sheets;
            ExcelWorksheets = ExcelSheets.Cast<Excel.Worksheet>();
            ExcelWorksheet = ExcelWorksheets.First(s => s.Name == Config.SheetName);
            ExcelRange = ExcelWorksheet.UsedRange;
        }

        public static void Main(string[] args)
        {
            // Generate and print SQL query to Debug window
            QueryGenerator();

            // Exit and release processes
            KillProcesses();
            Console.WriteLine("Finished killing Excel processes");

            // Trigger Garbage Collector twice
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.WriteLine("Finished collecting garbage the first time.");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.WriteLine("Finished collecting garbage the second time.");
        }

        private static void QueryGenerator()
        {
            string finalQuery = string.Empty;

            int rowCount = ExcelRange.Rows.Count;
            int colCount = ExcelRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                List<string> cellValues = RowDataExtractor(i, colCount, ExcelRange);

                string queryIndent = "    ";
                string queryStart = "Values (";
                string queryMiddle = string.Join(", ", cellValues);
                string queryEnd = ");";
                string queryTerminator = "\r\n";

                string insertQuery = QueryHelper.Queries[Config.SheetName] + queryTerminator;
                string valueQuery = string.Join(
                    string.Empty, new string[] { queryIndent, queryStart, queryMiddle, queryEnd, queryTerminator }
                );

                finalQuery += string.Join("", new string[] { insertQuery, valueQuery });
            }

            Console.WriteLine(finalQuery);
        }


        private static List<string> RowDataExtractor(int currentRow, int colCount, Excel.Range ExcelRange)
        {
            List<string> outputList = new();

            for (int j = 2; j <= colCount; j++)
            {
                var colTitle = ExcelRange.Cells[1, j].Value2;

                if (string.IsNullOrWhiteSpace(colTitle))
                {
                    break;
                }

                string? validVal;
                var cell = ExcelRange.Cells[currentRow, j];
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

        private static void KillProcesses()
        {
            // Rule of thumb for releasing com objects:
            //   never use two dots, all COM objects must be referenced and released individually
            //   ex: [somthing].[something].[something] is bad

            // Exit and release processes
            ExcelWorkbook.Close();
            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelRange);
            Marshal.ReleaseComObject(ExcelWorksheet);
            Marshal.ReleaseComObject(ExcelSheets);
            Marshal.ReleaseComObject(ExcelWorkbook);
            Marshal.ReleaseComObject(ExcelApp);

            Process[] prs = Process.GetProcessesByName("Excel");
            foreach (Process p in prs)
            {
                p.Kill();
            }
        }
    }
}
