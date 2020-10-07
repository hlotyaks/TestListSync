using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAO = Microsoft.Office.Interop.Access.Dao;
using System.Data.OleDb;
using System.IO;

namespace TestListSynchronizer
{
    /// <summary>
    /// 
    /// </summary>
    public class DatabaseSync
    {
        private List<int> _incomingSuiteIDs = new List<int>();
        private List<string> _errors = new List<string>();
        IDatabaseEngine _engine;
        ITestListSyncFactory _dbenginefactory;
        IRecordUpdater _recordUpdater;
        IJarvisWrapper _jarvis;
        string _dbName;
        string _dbTable;

        //
        // This link has useful information about install the db provider redistributable.
        // https://www.nicelabel.com/support/knowledge-base/article/using-excel-xlsx-and-access-accdb-data-source-in-office-365
        //

        /// <summary>
        /// 
        /// </summary>
        /// <param name="db"></param>
        /// <param name="table"></param>
        public DatabaseSync(string db, string table, ITestListSyncFactory factory, string jarvisApp)
        {
            _dbenginefactory = factory;
            _engine = _dbenginefactory.CreateDatabaseEngine();
            _recordUpdater = _dbenginefactory.CreateRecordUpdater();
            _jarvis = _dbenginefactory.CreateJarvisWrapper(jarvisApp);
            _dbName = db;
            _dbTable = table;
        }

        /// <summary>
        /// Will updated the database with the project data.  If parent data is provided that will also be updated
        /// Data is retrieved from Jarvis.
        /// </summary>
        /// <param name="project">Project name. Example: uflx2_PublicAPI</param>
        /// <param name="baseline">If not provided the latest baseline is used.  Other wise the data from Jarvis is for the provided baseline.</param>
        /// <param name="parentProject">Name of the parent project</param>
        /// <param name="parentProjectBaseline">Baseline to use for the parent data</param>
        public void UpdateDatabase(string project, string baseline, string parentProject, string parentProjectBaseline)
        {
            using (IDatabase db = _engine.Open(_dbName))
            {
                UpdateFromJarvis(db, _dbTable, project, baseline);
                
                if (parentProject != null)
                {
                    UpdateParentFromJarvis(db, _dbTable, parentProject, parentProjectBaseline);
                }

                UpdateDisabledTests(db, _dbTable);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public List<string> ErrorList
        {
            get => _errors;
        }

        /// <summary>
        /// 
        /// </summary>
        public bool IsErrors 
        {
            get => _errors.Count() != 0;
        }

        /// <summary>
        /// If there were any tests that were not run, but had previously been run then then are updated to 
        /// indicate they were "Not Run"
        /// </summary>
        /// <param name="db">database handle</param>
        /// <param name="accessTable">table in the database to update</param>
        private void UpdateDisabledTests(IDatabase db, string accessTable)
        {
            int currentTest = 0;
            int totalcount = db.TableSize(accessTable);

            string query = $"SELECT * FROM {accessTable}";

            using (IRecords record = db.OpenRecords(query))
            {
                Console.WriteLine($"\n\nUpdating status of tests not run");

                while (!record.EOF)
                {
                    currentTest++;
                    Console.Write($"\r Updating {currentTest} of {totalcount}");

                    int suiteID = record.GetSuiteID();

                    // the SuiteID in the database was not in the new data.
                    _recordUpdater.NotRunRecord(suiteID, record);

                    record.MoveNext();
                }
            }
        }

        /// <summary>
        /// Retrieve the data from Jarvis and use it to update the database
        /// </summary>
        /// <param name="db">database handle</param>
        /// <param name="accessTable">table in the database to update</param>
        /// <param name="project">Project name. Example: uflx2_PublicAPI</param>
        /// <param name="baseline">If not provided the latest baseline is used.  Other wise the data from Jarvis is for the provided baseline.</param>
        private void UpdateFromJarvis(IDatabase db, string accessTable, string project, string baseline)
        {
            // get project results from Jarvis
            List<SuiteResult> results = (baseline == null) ?
             _jarvis.FetchResults(project) :
             _jarvis.FetchResults(project, baseline);

            // get count of tests
            int totalTestCount = results.Count();
            int currentTest = 0;

            Console.WriteLine($"\n\nUpdating database with {project} test results");

            results.ForEach(item =>
            {
                currentTest++;
                Console.Write($"\r Updating {currentTest} of {totalTestCount}");

                // suiteid 
                int suiteID = item.SuiteID;

                // keep track of all the suite IDs we see in the data
                _recordUpdater.AddIncommingRecord(suiteID);

                // query for the suite id record in the access database
                string query = $"SELECT * FROM {accessTable} WHERE [Suite ID] = {suiteID}";

                using (IRecords dbRecords = db.OpenRecords(query))
                {
                    // EOF will = true if the suiteID was not found in the database.  This means it is a new record.
                    if (dbRecords.EOF)
                    {
                        // Add new record if it was not found
                        _recordUpdater.NewRecord(dbRecords, item);
                    }
                    else 
                    {
                        // Update the existing record
                        _recordUpdater.UpdateRecord(dbRecords, item);
                    }
                }

            });
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="db">database handle</param>
        /// <param name="dbTable">table in the database to update</param>
        /// <param name="parentProject">Name of the parent project</param>
        /// <param name="parentProjectBaseline">Baseline to use for the parent data</param>
        private void UpdateParentFromJarvis(IDatabase db, string dbTable, string parentProject, string parentProjectBaseline)
        {

            // get results for the parent project at a specific baseline
            List<SuiteResult> results = _jarvis.FetchResults(parentProject, parentProjectBaseline);

            // get count of tests
            int totalTestCount = results.Count();
            int currentTest = 0;

            Console.WriteLine($"\n\nUpdating database with parent results {parentProject}@{parentProjectBaseline}");

            results.ForEach(item =>
            {
                currentTest++;
                Console.Write($"\r Updating {currentTest} of {totalTestCount}");

                // suiteid 
                int suiteID = item.SuiteID;

                // query for the suite id record in the access database
                string query = $"SELECT * FROM {dbTable} WHERE [Suite ID] = {suiteID}";

                using (IRecords dbRecords = db.OpenRecords(query))
                {

                    // If we find a test in the parent data that is not in the database then we are probably 
                    // using the wrong date for the parent data since the rebase should have brought over new tests.
                    if (dbRecords.EOF)
                    {
                        _errors.Add($"Suite ID [{suiteID}] not found in database. Possible use of incorrect parent branch result data");
                    }
                    else 
                    {
                        _recordUpdater.UpdateParentRecord(dbRecords, item);
                    }
                }
            });
        }
    }
}
