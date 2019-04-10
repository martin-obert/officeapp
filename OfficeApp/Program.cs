using System;
using System.Collections.Generic;
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
            var filename = ConfigurationManager.AppSettings["filename"];
            var directory = ConfigurationManager.AppSettings["directory"];

            Console.WriteLine("Starting headless application mode.\nLooking for files: \"" + filename + "\"");

            if (!Directory.Exists(directory))
                Console.WriteLine("Non existing directory: " + directory);



            if (!(ConfigurationManager.GetSection("ExcelFormattersSection") is ExcelFormattersSection section))
                throw new NullReferenceException("Configuration section for formatting was not found.");

            Console.WriteLine("Reading formatters configuration:");

            var formatters = section.Formatters.Cast<ExcelFormatterElement>();

            var regexPattern = new StringBuilder();

            foreach (var formatter in formatters)
            {
                if (string.IsNullOrEmpty(formatter.Keyword) || string.IsNullOrEmpty(formatter.Replacement))
                    continue;

                Console.WriteLine("Keyword: \"" + formatter.Keyword + "\" will be replaced with: \"" +
                                  formatter.Replacement + "\"");

                if (regexPattern.Length > 0)
                {
                    regexPattern.Append("|");
                }

                regexPattern.AppendFormat("(?<g{0}>{1})", formatter.Id, formatter.Keyword);
            }
            Console.Write("Regex: " + regexPattern);
            var regex = new Regex(regexPattern.ToString());

            var cellsRange = ConfigurationManager.AppSettings["range"];

            if (string.IsNullOrEmpty(cellsRange))
            {
                Console.WriteLine(
                    "No range was specified, please add or update value in App.config -> appSettings -> range");
                return;
            }

            var maxFileId = FindMaxIdFromFolder(directory, filename);

            ProcessFile(Path.Combine(directory, $"{filename}{maxFileId:000}"), cellsRange, regex, formatters);

            Console.WriteLine("Cells range is set to: " + cellsRange);


        }

        private static Regex _numberRegex = new Regex(@"(?<id>\d{3}$)");


        private static void ProcessFile(string filepath, string cellsRange, Regex regex, IEnumerable<ExcelFormatterElement> formatters)
        {
            var excel = new Application();
            Workbooks wkbks = null;

            wkbks = excel.Workbooks;

            Workbook wkbk = null;
            wkbk = wkbks.Open(Path.Combine(Directory.GetCurrentDirectory(), filepath + ".xls"));

            excel.Visible = false;

            excel.DisplayAlerts = false;

            try
            {
                var wb = excel.ActiveWorkbook;

                Console.WriteLine(wb.Worksheets.Count);

                var ws = wb.Worksheets[1] as Worksheet;

                foreach (var excelFormatterElement in formatters.Where(x => !string.IsNullOrEmpty(x.Range)))
                {
                    if (!string.IsNullOrEmpty(excelFormatterElement.NumberFormat))
                        ws.Range[excelFormatterElement.Range].EntireColumn.NumberFormat = excelFormatterElement.NumberFormat;
                }



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
                            DateTime date;
                            TimeSpan timespan;
                            if (DateTime.TryParse(val.ToString(), out date) && !TimeSpan.TryParse(val.ToString(), out timespan) && date > new DateTime(2002))
                            {
                                currentRange.Value = date.ToString("dd.MM.yyyy");
                            }

                            continue;
                        }

                        

                        Console.WriteLine("Processing cell: COL " + currentRange.Column + " ROW " + currentRange.Row);


                        string value = currentRange.Value.ToString();

                        value = regex.Replace(value, match =>
                        {
                            return formatters.First(x => x.Keyword == match.Value).Replacement;
                        });

                        Console.WriteLine("Replacing: " + value + " -> " + currentRange.Value.ToString());


                        currentRange.Value = value;

                    }
                }


                wb.Save();
                //wb.Save(filepath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing,
                //    Type.Missing,
                //    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                //    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();

                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wkbk);
                Marshal.ReleaseComObject(wkbks);
                var newFilename = Path.Combine(Directory.GetCurrentDirectory(), filepath + ".xlsx");
                if(File.Exists(newFilename))
                    File.Delete(newFilename);
                File.Move(Path.Combine(Directory.GetCurrentDirectory(), filepath + ".xls"), newFilename);


                Console.WriteLine("Saving to: " + filepath);

                Console.WriteLine("Finished succesfully");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Trace.WriteLine(e.ToString());
                Console.WriteLine("Finished with errors. Please read application.log at " +
                                  Directory.GetCurrentDirectory());
            }
            finally
            {
                excel.Quit();

                GC.Collect();

                GC.WaitForPendingFinalizers();
                KillExcel();


            }
        }

        private static void KillExcel()
        {
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0)
                {
                    PK.Kill();
                }
            }
        }

        private static void SaveMaxId(int value)
        {
            var conf = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            conf.AppSettings.Settings["maxId"].Value = value.ToString("000");
        }

        private static int LoadMaxId()
        {
            int.TryParse(ConfigurationManager.AppSettings["maxId"], out var maxId);
            return maxId;
        }
        private static int FindMaxIdFromFolder(string searchDirectory, string filenameRaw)
        {
            if (string.IsNullOrEmpty(searchDirectory))
            {
                Console.WriteLine("Directory with excel files not provided.");
                return 0;
            }

            if (!Directory.Exists(searchDirectory))
            {
                Console.WriteLine("Directory with excel files was not found.");
                return 0;
            }

            var files = Directory.GetFiles(searchDirectory);

            var maxId = 0;


            foreach (var file in files)
            {
                if (filenameRaw != null && !Regex.IsMatch(file, filenameRaw))
                    continue;

                var regex = _numberRegex.Match(Path.GetFileNameWithoutExtension(file));

                if (regex.Length > 0 && regex.Groups["id"] != null)
                {
                    var idRaw = regex.Groups["id"].Value;

                    if (int.TryParse(idRaw, out var temp))
                    {
                        if (temp > maxId)
                        {
                            maxId = temp;
                        }
                    }
                }
            }

            return maxId;

        }
    }
}