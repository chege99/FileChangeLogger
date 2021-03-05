using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using IronXL;

namespace FileChangeLogger
{
    class Program
    {
        static string workbook_full_path = @"C:\Users\Kobe\Desktop\DAPROIM\LogFile.xlsx";
        static void Main(string[] args)
        {
            var parent_directory = @"C:\Users\Kobe\Desktop\JUNK";
            var watcher = new FileSystemWatcher(parent_directory);

            //Changes to be monitored
            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;
           //Events To be traced
            watcher.Created += OnCreated;
            watcher.Error += OnError;

            watcher.Filter = "*.*";

            watcher.IncludeSubdirectories = true;
            watcher.EnableRaisingEvents = true;

            Console.WriteLine("Press Enter to Exit...");
            Console.ReadLine();

        }

        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            string filePath= $"Created: {e.FullPath}";
            var timeStamp=DateTime.Now;
            ExportToExcel(filePath, timeStamp.ToString());
        }

        private static void OnError(object sender, ErrorEventArgs e) =>
            PrintException(e.GetException()
        );

        private static void PrintException(Exception ex)
        {
            if (ex != null)
            {
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine("Stacktrace:");
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine();
                PrintException(ex.InnerException);
            }
        }
        private static void ExportToExcel(string Path, string Timestamp)
        {
            Console.WriteLine(Path);
            Console.WriteLine(Timestamp);
            //Load Work Book
            WorkBook custom_workbook = WorkBook.Load(workbook_full_path);

            string date_today = DateTime.Today.ToShortDateString();
            string custom_worksheet = date_today.Replace("/", "_");
            //Check If Today's Worksheet Exists
            WorkSheet todays_sheet;

            if (custom_workbook.GetWorkSheet(custom_worksheet) == null)
            {
                todays_sheet = custom_workbook.CreateWorkSheet(custom_worksheet);
                todays_sheet["A1"].Value = "Path";
                todays_sheet["B1"].Value = "Date Create";
                custom_workbook.Save();
            }
            else
            {
                todays_sheet = custom_workbook.GetWorkSheet(custom_worksheet);
            }

            //Get Last Row info
            int row_count = todays_sheet.Rows.Count();
            int row_number = row_count + 1;

            //Sheet Data
            todays_sheet["A" + row_number].Value = Path;
            todays_sheet["B" + row_number].Value = Timestamp;

            custom_workbook.Save();

        }
    }
}
