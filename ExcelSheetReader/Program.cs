#pragma warning disable S1075, S1215, CA1416

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelSheetReader.Helpers;
using ExcelSheetReader.Settings;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace ExcelSheetReader
{
    public static class Program
    {
        public static Excel.Application ExcelApp { get; private set; }
        public static Excel.Workbook ExcelWorkbook { get; private set; }
        public static Excel.Sheets ExcelSheets { get; private set; }
        public static IEnumerable<Excel.Worksheet> ExcelWorksheets { get; private set; }

        static Program()
        {
            // Create COM Objects. Create a COM object for everything that is referenced
            ExcelApp = new Excel.Application();
            ExcelWorkbook = ExcelApp.Workbooks.Open(Config.FilePath);
            ExcelSheets = ExcelWorkbook.Sheets;
            ExcelWorksheets = ExcelSheets.Cast<Excel.Worksheet>();
        }

        public static void Main(string[] args)
        {
            foreach (SheetInfo sheetInfo in Config.WorkBookInfo)
            {
                // Generate and print SQL query to Debug window
                QueryGenerator(sheetInfo);
            }
            
            // Exit and release processes
            KillAllExcelProcesses();
            Console.WriteLine("Finished killing Excel processes");

            // Trigger Garbage Collector twice
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.WriteLine("Finished collecting garbage the first time.");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Console.WriteLine("Finished collecting garbage the second time.");
        }

        private static void QueryGenerator(SheetInfo sheetInfo)
        {
            StringBuilder queryBuilder = new(); 
            int startingCol = sheetInfo.IsJunctionTable ? 1 : 2;

            Excel._Worksheet excelWorksheet = ExcelWorksheets.First(s => s.Name == sheetInfo.SheetName);
            Excel.Range excelRange = excelWorksheet.UsedRange;
            int rowCount = excelRange.Rows.Count;
            int colCount = excelRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                List<string> cellValues = RowDataExtractor(excelRange, i, colCount, startingCol);

                string queryIndent = "    ";
                string queryStart = "Values (";
                string queryMiddle = string.Join(", ", cellValues);
                string queryEnd = ");";
                string queryTerminator = "\r\n";

                List<string> valueQueryList = new()
                {
                    queryIndent, queryStart, queryMiddle, queryEnd, queryTerminator
                };
                List<string> setVarQueryList = new()
                {
                    queryIndent,
                    QueryHelper.GenerateQuery_SetTemporaryVariableToLastInsertId(
                        sheetInfo.TempVarName, i-1
                    ),
                    queryTerminator
                };

                string insertQuery = QueryHelper.Queries[excelWorksheet.Name] + queryTerminator;
                string valueQuery = string.Join(string.Empty, valueQueryList);
                string setVarQuery = sheetInfo.IsJunctionTable
                    ? string.Empty
                    : string.Join(string.Empty, setVarQueryList);

                queryBuilder.Append(insertQuery + valueQuery + setVarQuery);
            }

            ReleaseComObjects(ref excelRange, ref excelWorksheet);
            Console.WriteLine(queryBuilder.ToString());
        }


        private static List<string> RowDataExtractor(
            Excel.Range ExcelRange, int currentRow, int colCount, int startingCol
        )
        {
            List<string> outputList = new();

            for (int j = startingCol; j <= colCount; j++)
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

                validVal = cellValIsCode 
                    ? $"{CodeHelper.Codes[cellVal]}" 
                    : FormatNonCodeData(cellVal ?? "null");

                outputList.Add(validVal);
            }

            return outputList;
        }

        private static string FormatNonCodeData<T>(T cellVal)
        {
            // Format null values
            List<string> nullList = new() { "null", "NULL" };
            string output = (!nullList.Contains($"{cellVal}")) ? $"'{cellVal}'" : "null";

            bool cellValIsTempVarVal = output.Contains("'@");
            output = cellValIsTempVarVal ? output.Trim('\'') : output;

            // Format dates
            Regex rgx = new Regex("\\d\\d/\\d\\d/\\d\\d\\d\\d", RegexOptions.IgnoreCase);
            if (rgx.IsMatch(output))
            {
                Match match = rgx.Match(output);
                DateTime.TryParse(match.Value + " 00:00:00", out DateTime parsedDate);
                output = rgx.Replace(output, parsedDate.ToString("yyyy-MM-dd HH-mm-ss"));
            }

            return output;
        }

        private static void ReleaseComObjects(
            ref Excel.Range excelRange, ref Excel._Worksheet excelWorksheet
        )
        {
            // Release range and worksheet objects to free up resources for the next sheet
            Marshal.ReleaseComObject(excelRange);
            Marshal.ReleaseComObject(excelWorksheet);
        }

        private static void KillAllExcelProcesses()
        {
            // Rule of thumb for releasing com objects:
            //   never use two dots, all COM objects must be referenced and released individually
            //   ex: [somthing].[something].[something] is bad

            // Exit and release remaining processes
            ExcelWorkbook.Close();
            ExcelApp.Quit();
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
