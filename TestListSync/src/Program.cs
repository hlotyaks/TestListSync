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
            [Option("project", Required = true, HelpText = "Project name")]
            public string Project { get; set; }

            [Option("baseline", Required = false, HelpText = "Baseline for test results")]
            public string ProjectBaseline { get; set; }

            [Option("parentProject", Required = false, HelpText = "Project name for parent branch")]
            public string ProjectParent { get; set; }

            [Option("parentBaseline", Required = false, HelpText = "Baseline for results from parent branch")]
            public string ProjectParentBaseline { get; set; }

            [Option("jarvisConsolePath", Required = false, HelpText = "Path to location of Jarvis.exe.  Default is ./")]
            public string JarvisConsolePath { get; set; }

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
            string jarvisPath = null;
            string jarvisExe = "Jarvis.exe";
            string project = null;
            string baseline = null;
            string projectParent = null;
            string projectParentBaseline = null;
            string dbFile=null;
            string dbTable = null; ;
            bool IllegalCommands = false;

            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o =>
                {
                    if (o.Help)
                    {
                        ShowHelp();
                        return;
                    }

                    if (o.JarvisConsolePath == null)
                    {
                        jarvisPath = $"{System.IO.Directory.GetCurrentDirectory()}\\{jarvisExe}";
                    }
                    else
                    {
                        jarvisPath = $"{o.JarvisConsolePath}\\{jarvisExe}";
                    }

                    //
                    // Get the database file name
                    //
                    if (o.DatabaseFile == null)
                    {
                        ShowHelp();
                        IllegalCommands = true;
                    }
                    else
                    {
                        dbFile = o.DatabaseFile;
                    }

                    //
                    // Get the database table name
                    //
                    if (o.DatabaseTable == null)
                    {
                        ShowHelp();
                        IllegalCommands = true;
                    }
                    else
                    {
                        dbTable = o.DatabaseTable;
                    }

                    //
                    // Get the project name
                    //
                    if (o.Project == null)
                    {
                        ShowHelp();
                        IllegalCommands = true;
                    }
                    else
                    {
                        project = o.Project;
                    }

                    //
                    // Get th ebaseline to us for the project.  Is optional.
                    //
                    if (o.ProjectBaseline == null)
                    {
                        baseline = null;
                    }
                    else
                    {
                        baseline = o.ProjectBaseline;
                    }

                    // if parent project is provided then the parent baseline is also required
                    if(o.ProjectParent != null && o.ProjectParentBaseline != null)
                    {
                        projectParent = o.ProjectParent;
                        projectParentBaseline = o.ProjectParentBaseline;
                    }
                    else if (o.ProjectParent == null && o.ProjectParentBaseline == null)
                    {
                        // if both are null then nothing to do.  This is legal since both are optional
                    }
                    else
                    {
                        // only one option was provided  
                        ShowHelp();
                        IllegalCommands = true;
                    }

                    ITestListSyncFactory factory = new TestListSyncFactory();
                    DatabaseSync dbsync = new TestListSynchronizer.DatabaseSync(dbFile, dbTable, factory);

                    try
                    {
                        if (!System.IO.File.Exists(jarvisPath))
                        {
                            throw new Exceptions.JarvisNotFoundException($"{jarvisPath} not found.");
                        }

                        if (!IllegalCommands)
                        {
                            dbsync.UpdateDatabase(project, baseline, projectParent, projectParentBaseline);
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
                    catch (Exceptions.JarvisNotFoundException e)
                    {
                        Console.WriteLine($"Excpetion: {e.Message}");
                    }
                    finally
                    {
                        if (!dbsync.IsErrors)
                        {
                            Console.WriteLine("");
                            Console.WriteLine("Warnings:");
                            dbsync.ErrorList.ForEach(s => Console.WriteLine(s));
                        }
                    }

                });
        }

        private static void ShowHelp()
        {
            Console.WriteLine("TestListSync");
            Console.WriteLine("");
            Console.WriteLine("A utility to update a sharepoint list containing tests. Inputs include project and baseline information for retrieving data from Jarvis.");
            Console.WriteLine("");
            Console.WriteLine("-project             : Project name. Example: uflx2_PublicAPI");
            Console.WriteLine("-baseline            : Baseline to use for test data. Default is latest baseline.");
            Console.WriteLine("-parentProject       : Project name for the parent branch.");
            Console.WriteLine("-parentBaseline      : Baseline on parent branch to use for parent test data. RTequired if parent branch is supplied.");
            Console.WriteLine("-jarvisConsolePath   : Path to Jarvis.exe. Default is ./");
            Console.WriteLine("-database            : Path to database file that is synced with the Sharepoint site");
            Console.WriteLine("-table               : Name of database table that will be updated with the Excel data");
            Console.WriteLine("");
            Console.WriteLine(@"Example: TestListSyc --project uflx2_PublicAPI --database C:\tmp\database.accdb -table table1");
            Console.WriteLine("");

        }
    }
}
