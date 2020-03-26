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

            [Option('p', "parent", Required = false, HelpText = "Parent Files containing test results from last RI")]
            public IEnumerable<string> ParentrTestFiles { get; set; }

            [Option('d', "database", Required = true, HelpText = "Database file name containing test results")]
            public string DatabaseFile { get; set; }

            [Option('t', "table", Required = true, HelpText = "Table in database containing test results")]
            public string DatabaseTable { get; set; }

            [Option('h', "help", HelpText = "Show help")]
            public bool Help { get; set; }
        }

        static void Main(string[] args)
        {
            List<string> InputFiles = new List<string>();
            List<string> ParentFiles = new List<string>();
            string dbFile=null;
            string dbTable = null; ;
            bool IllegalCommands = false;

            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       if(o.Help)
                       {
                           ShowHelp();
                           return;
                       }

                       // Check for the test files
                       if (o.TestFiles.Count() == 0)
                       {
                           ShowHelp();
                           IllegalCommands = true;
                       }
                       else
                       {
                           InputFiles.AddRange(o.TestFiles);
                       }

                       // Check for the parent test files.  These are optional
                       if (o.ParentrTestFiles.Count() != 0)
                       {
                           ParentFiles.AddRange(o.ParentrTestFiles);
                           // Has to be 2 files. ASRT and BFR
                           if (ParentFiles.Count != 2)
                           {
                               ShowHelp();
                               IllegalCommands = true;
                           }
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

                       DatabaseSync dbsync = new TestListSynchronizer.DatabaseSync(dbFile, dbTable);

                       try
                       {
                           if (!IllegalCommands)
                           {                       
                               dbsync.UpdateDatabase(InputFiles[0], InputFiles[1], ParentFiles[0], ParentFiles[1]);
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
                       catch (Exceptions.DatabaseOpenException e)
                       {
                           Console.WriteLine($"Exception: Error opening database {e.Message}. Verify it is not currently open.");
                       }
                       finally
                       {
                           if (!dbsync.IsErrors)
                           {
                               Console.WriteLine("");
                               Console.WriteLine("Errors:");
                               dbsync.ErrorList.ForEach(s => Console.WriteLine(s));
                           }
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
            Console.WriteLine("-p : Path to files containing parent branch test results.  This should be results from the point of the last integration.");
            Console.WriteLine("");
            Console.WriteLine(@"Example: TestListSyc -f C:\tmp\asrt.xlsx C:\tmp\bfr.xlsx -d C:\tmp\database.accdb -t Table1");
            Console.WriteLine(@"Example: TestListSyc -f C:\tmp\asrt.xlsx C:\tmp\bfr.xlsx -p C:\tmp\parent-asrt.xlsx C:\tmp\parnet-bfr.xlsx -d C:\tmp\database.accdb -t Table1");
            Console.WriteLine("");

        }
    }
}
