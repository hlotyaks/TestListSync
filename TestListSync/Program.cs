using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using TestListSynchronizer;
using Exceptions = TestListSynchronizer.Exceptions;

namespace TestListSync
{
    class Program
    {
        public class Options
        {
            [Option('f', "files", Required = true, HelpText = "Files containing latest test results")]
            public IEnumerable<string> TestFiles { get; set; }

            [Option('d', "database", Required = true, HelpText = "Database file name containing test results")]
            public string DatabaseFile { get; set; }

            [Option('t', "table", Required = true, HelpText = "Table in database containing test results")]
            public string DatabaseTable { get; set; }
        }

        static void Main(string[] args)
        {
            List<string> InputFiles = new List<string>();
            string dbFile=null;
            string dbTable = null; ;
            bool IllegalCommands = false;

            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       if (o.TestFiles.Count() == 0)
                       {
                           ShowHelp();
                           IllegalCommands = true;
                       }
                       else
                       {
                           InputFiles.AddRange(o.TestFiles);
                       }

                       if (o.DatabaseFile == null)
                       {
                           ShowHelp();
                           IllegalCommands = true;
                       }
                       else
                       {
                           dbFile = o.DatabaseFile;
                       }

                       if(o.DatabaseTable == null)
                       {
                           ShowHelp();
                           IllegalCommands = true;
                       }
                       else
                       {
                           dbTable = o.DatabaseTable;
                       }

                       try
                       {
                           if (!IllegalCommands)
                           {
                               DatabaseSync dbsync = new TestListSynchronizer.DatabaseSync(dbFile, dbTable);
                               dbsync.UpdateDatabase(InputFiles[0], InputFiles[1]);
                           }
                       }
                       catch (Exceptions.ExcelSheetCountException e)
                       {
                           Console.WriteLine($"Excpetion: Illegal number of sheets in spreadsheet {e.Message}. Must be 1.");
                       }
                       catch (Exceptions.ExcelTestCountException e)
                       {
                           Console.WriteLine($"Exception: No tests in spreadsheet {e.Message}.");
                       }
                   });
        }

        private static void ShowHelp()
        {
            Console.WriteLine("TestListSync");
            Console.WriteLine("");
            Console.WriteLine("A utility to update a sharepoint list containing tests. Inputs include Excel spreadsheets exported from Jarvis");
            Console.WriteLine("");
            Console.WriteLine("-f : Path to Excel files containing latest test data");
            Console.WriteLine("-d : Path to database file that is synced with the Sharepoint site");
            Console.WriteLine("-t : Name of database table that will be updated with the Excel data");
            Console.WriteLine("");
            Console.WriteLine(@"Example: TestListSyc -f C:\tmp\asrt.xlsx C:\tmp\bfr.xlsx -d C:\tmp\database.accdb -t Table1");
            Console.WriteLine("");

        }
    }
}
