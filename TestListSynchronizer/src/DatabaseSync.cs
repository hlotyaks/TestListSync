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
        private List<string> errors = new List<string>();
        IDatabaseEngine _engine;
        IDatabaseEngineFactory _dbenginefactory;
        IRecordUpdater _recordUpdater;
        string dbName;
        string dbTable;

        //
        // This link has useful information about install the db provider redistributable.
        // https://www.nicelabel.com/support/knowledge-base/article/using-excel-xlsx-and-access-accdb-data-source-in-office-365
        //

        /// <summary>
        /// 
        /// </summary>
        /// <param name="db"></param>
        /// <param name="table"></param>
        public DatabaseSync(string db, string table, IDatabaseEngineFactory factory)
        {
            _dbenginefactory = factory;
            _engine = _dbenginefactory.CreateDatabaseEngine();
            _recordUpdater = _dbenginefactory.CreateRecordUpdater();
            dbName = db;
            dbTable = table;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filelist"></param>
        /// <param name="parentfilelist"></param>
        public void UpdateDatabase(List<string> filelist)
        {
            using (IDatabase db = _engine.Open(dbName))
            {
                filelist.ForEach(item => UpdateFromFile(db, dbTable, item));
                UpdateDisabledTests(db, dbTable);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filelist"></param>
        /// <param name="parentfilelist"></param>
        public void UpdateDatabase(List<string> filelist, List<string> parentfilelist)
        {
            using (IDatabase db = _engine.Open(dbName))
            {
                filelist.ForEach(item => UpdateFromFile(db, dbTable, item));

                // parent data files
                filelist.ForEach(item => UpdateFromParentFile(db, dbTable, item));
                
                UpdateDisabledTests(db, dbTable);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public List<string> ErrorList
        {
            get => errors;
        }

        /// <summary>
        /// 
        /// </summary>
        public bool IsErrors 
        {
            get => errors.Count() != 0;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="db"></param>
        /// <param name="accessTable"></param>
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
        /// 
        /// </summary>
        /// <param name="db"></param>
        /// <param name="accessTable"></param>
        /// <param name="xlsxFile"></param>
        private void UpdateFromFile(IDatabase db, string accessTable, string xlsxFile)
        {
            // get count of tests
            int totalTestCount = TestCount(xlsxFile);
            int currentTest = 0;

            using (IDatabase xldb = _engine.Open(xlsxFile, "Excel 12.0 Xml;HDR=YES;"))
            {
                if (xldb.TableCount != 1)
                {
                    throw new Exceptions.ExcelSheetCountException(xlsxFile);
                }

                using (IRecords xlRecord = xldb.OpenRecords(xldb.TableName(0)))
                {
                    Console.WriteLine($"\n\nUpdating database with {xlsxFile}");

                    while (!xlRecord.EOF)
                    {
                        currentTest++;
                        Console.Write($"\r Updating {currentTest} of {totalTestCount}");

                        // suiteid comes from excel as a double
                        int suiteID = xlRecord.GetSuiteID();

                        // keep track of all the suite IDs we see in the data
                        _recordUpdater.AddIncommingRecord(suiteID);

                        // query for the suite id record in the access database
                        string query = $"SELECT * FROM {accessTable} WHERE [Suite ID] = {suiteID}";

                        using (IRecords dbRecords = db.OpenRecords(query))
                        {
                            // Check for existence of suite id in the database. Add new record if new
                            if (dbRecords.EOF)
                            {
                                _recordUpdater.NewRecord(dbRecords, xlRecord);
                                xlRecord.MoveNext();
                                continue;
                            }

                            // Update the existing record
                            _recordUpdater.UpdateRecord(dbRecords, xlRecord);
                           
                            xlRecord.MoveNext();
                        }
                    }
                }
            }
        }
    
        /// <summary>
        /// 
        /// </summary>
        /// <param name="db"></param>
        /// <param name="accessTable"></param>
        /// <param name="xlsxFile"></param>
        private void UpdateFromParentFile(IDatabase db, string accessTable, string xlsxFile)
        {
            // get count of tests
            int totalTestCount = TestCount(xlsxFile);
            int currentTest = 0;

            using (IDatabase xldb = _engine.Open(xlsxFile, "Excel 12.0 Xml;HDR=YES;"))
            {
                if (xldb.TableCount != 1)
                {
                    throw new Exceptions.ExcelSheetCountException(xlsxFile);
                }

                using (IRecords xlRecords = xldb.OpenRecords(xldb.TableName(0)))
                {
                    Console.WriteLine($"\n\nUpdating database with {xlsxFile}");

                    while (!xlRecords.EOF)
                    {
                        currentTest++;
                        Console.Write($"\r Updating {currentTest} of {totalTestCount}");

                        // suiteid comes from excel as a double
                        int suiteID = xlRecords.GetSuiteID();

                        // query for the suite id record in the access database
                        string query = $"SELECT * FROM {accessTable} WHERE [Suite ID] = {suiteID}";

                        using (IRecords dbRecords = db.OpenRecords(query))
                        {

                            // If we find a test in the parent data that is not in the database then we are probably 
                            // using the wrong date for the parent data since the rebase should have brought over new tests.
                            if (dbRecords.EOF)
                            {
                                errors.Add($"Suite ID [{suiteID}] not found in database. Possible use of incorrect parent branch result data");
                                xlRecords.MoveNext();
                                continue;
                            }

                            _recordUpdater.UpdateParentRecord(dbRecords, xlRecords);

                            xlRecords.MoveNext();
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xlsxFile"></param>
        /// <returns></returns>
        public int TestCount(string xlsxFile)
        {
            IDatabaseEngine engine = _dbenginefactory.CreateDatabaseEngine();
            int testCount;

            using (IDatabase db = engine.Open(xlsxFile, "Excel 12.0 Xml;HDR=YES;"))
            {
                if (db.TableCount != 1)
                {
                    throw new Exceptions.ExcelSheetCountException(xlsxFile);
                }

                string query = $"SELECT Count(*) as [CountOfRows] FROM [{db.TableName(0)}]";

                using (IRecords r = db.OpenRecords(query))
                {
                    testCount = (int)r.GetFieldValue("CountOfRows");
                }
            }

            if (testCount < 0)
            {
                throw new Exceptions.ExcelTestCountException(xlsxFile);
            }

            return testCount;
        }
    
      
    }
}
