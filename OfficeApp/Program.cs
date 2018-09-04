using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using OfficeApp.ConfigSections;

namespace OfficeApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var filepath = ConfigurationManager.AppSettings["filepath"];
            Console.WriteLine("Starting headless application mode.\nLooking for file: \"" + filepath + "\"");

            if (!string.IsNullOrEmpty(filepath) && File.Exists(filepath))
                Console.WriteLine("File - FOUND");
            else
            {
                Console.WriteLine("File - NOT FOUND, please add or update filepath in App.config -> appSettings -> filepath");
                return;
            }

            if (!(ConfigurationManager.GetSection("ExcelFormattersSection") is ExcelFormattersSection section))
                throw new NullReferenceException("Configuration section for formatting was not found.");

            Console.WriteLine("Reading formatters configuration:");

            var formatters = section.Formatters.Cast<ExcelFormatterElement>().Select((item, index) => new { index = index + 1, item });

            var regexPattern = new StringBuilder();

            foreach (var formatter in formatters)
            {
                Console.WriteLine("Keyword: \"" + formatter.item.Keyword + "\" will be replaced with: \"" + formatter.item.Replacement + "\"");

                if (regexPattern.Length > 0)
                {
                    regexPattern.Append("|");
                }

                regexPattern.AppendFormat("(?<{0}>{1})", formatter.index, formatter.item.Keyword);
            }

            var regex = new Regex(regexPattern.ToString());

            var cellsRange = ConfigurationManager.AppSettings["range"];

            if (string.IsNullOrEmpty(cellsRange))
            {
                Console.WriteLine("No range was specified, please add or update value in App.config -> appSettings -> range");
                return;
            }

            Console.WriteLine("Cells range is set to: " + cellsRange);

            var excel = new Application();
            Workbooks wkbks = null;

            wkbks = excel.Workbooks;

            Workbook wkbk = null;
            wkbk = wkbks.Open(Path.Combine(Directory.GetCurrentDirectory(), filepath));
            //wkbk.Activate();

            excel.Visible = false;
            excel.DisplayAlerts = false;
            try
            {

                var wb = excel.ActiveWorkbook;

                Console.WriteLine(wb.Worksheets.Count);

                var ws = wb.Worksheets[1] as Worksheet;

                var range = ws.Range[cellsRange];

                var rows = range.Rows.Count;

                var cols = range.Columns.Count;

                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= cols; j++)
                    {
                        var currentRange = (Range)range.Cells[i, j];

                        var val = currentRange.Value;

                        if (val == null)
                        {
                            continue;
                        }

                        var matches = regex.Matches(currentRange.Value.ToString()) as MatchCollection;

                        if (matches == null || matches.Count == 0)
                        {
                            continue;
                        }

                        Console.WriteLine("Processing cell: COL " + currentRange.Column + " ROW " + currentRange.Row);



                        string value = currentRange.Value.ToString();

                        value = regex.Replace(value, match =>
                        {
                            return formatters.First(x => x.item.Keyword == match.Value).item.Replacement;
                        });

                        Console.WriteLine("Replacing: " + value + " -> " + currentRange.Value.ToString());

                        currentRange.Value = value;

                    }
                }

                Console.WriteLine("Saving to: " + filepath);
                wb.SaveAs(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();

                Marshal.ReleaseComObject(wb);

                Console.WriteLine("Finished succesfully");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Trace.WriteLine(e.ToString());
                Console.WriteLine("Finished with errors. Please read application.log at " + Directory.GetCurrentDirectory());
            }
            finally
            {
                GC.Collect();

                GC.WaitForPendingFinalizers();

                excel.Quit();
            }
        }
    }
}
