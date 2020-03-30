using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Access.Dao;
using System.Data.OleDb;
using System.IO;
//using ADOX;

namespace TestListSynchronizer
{
    public class DatabaseSync
    {
        private const string SUITEID = "Suite ID";
        private List<int> newDataIDs = new List<int>();
        private List<string> errors = new List<string>();
        Database dbAcess;
        DBEngine dbEn;
        string dbName;
        string dbTable;

        //
        // This link has useful information about install the db provider redistributable.
        // https://www.nicelabel.com/support/knowledge-base/article/using-excel-xlsx-and-access-accdb-data-source-in-office-365
        //

        public DatabaseSync(string db, string table)
        {
            dbEn = new DBEngineClass();
            dbName = db;
            dbTable = table;
        }

        public void UpdateDatabase(List<string> filelist, List<string> parentfilelist)
        {
            try
            {
                dbAcess = dbEn.OpenDatabase(dbName);
            }
            catch (Exception)
            {
                throw new Exceptions.DatabaseOpenException(dbName);
            }
            // Refresh the data. will pull from sharepoint.
            dbAcess.TableDefs.Refresh();

            UpdateFromFile(dbAcess, dbTable, filelist[0]);
            UpdateFromFile(dbAcess, dbTable, filelist[1]);
            if (parentfilelist.Count() == 2)
            {
                UpdateFromParentFile(dbAcess, dbTable, parentfilelist[0]);
                UpdateFromParentFile(dbAcess, dbTable, parentfilelist[1]);
            }
            UpdateDisabledTests(dbAcess, dbTable);

            dbAcess.Close();
        }

        public List<string> ErrorList
        {
            get => errors;
        }

        public bool IsErrors 
        {
            get => errors.Count() != 0;
        }

        private void UpdateDisabledTests(Database dbAcess, string accessTable)
        {
            int currentcount = 0;
            string query = $"SELECT Count(*) as [CountOfRows] FROM {accessTable}";
            Recordset recordsCount = dbAcess.OpenRecordset(query, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);
            int totalcount = (int)recordsCount.Fields["CountOfRows"].Value;
            recordsCount.Close();

            query = $"SELECT * FROM {accessTable}";
            Recordset rs = dbAcess.OpenRecordset(query, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);

            Console.WriteLine($"\n\nUpdating status of tests not run");

            while (!rs.EOF)
            {
                currentcount++;
                Console.Write($"\r Updating {currentcount} of {totalcount}");

                object oSuiteID = rs.Fields[SUITEID].Value;
                int suiteID = (int)((double)oSuiteID);

                // the SuiteID in the database was not in the new data.
                if (!newDataIDs.Any(i => i == suiteID))
                {
                    rs.Edit();
                    rs.Fields["Result"].Value = "Not Run";
                    rs.Update(1, false);
                }

                rs.MoveNext();
            }

            rs.Close();
        }

        private void UpdateFromFile(Database db, string accessTable, string xlsxFile)
        {
            // get count of tests
            int totalcount = TestCount(xlsxFile);
            int currentcount = 0;
            
            // really need some sort of using block around the excel access...
            Recordset recordsExcel = OpenExcelRecords(xlsxFile);

            Console.WriteLine($"\n\nUpdating database with {xlsxFile}");

            while (!recordsExcel.EOF)
            {
                currentcount++;
                Console.Write($"\r Updating {currentcount} of {totalcount}");

                object suiteID = recordsExcel.Fields[SUITEID].Value; // suiteid comes from excel as a double
                newDataIDs.Add((int)((double)suiteID));

                string suitename = recordsExcel.Fields["Name"].Value as string; // Name stands for 'suite name'
                
                // query for the suite id record in the access database
                string query = $"SELECT * FROM {accessTable} WHERE [Suite ID] = {suiteID}";
                Recordset rs = db.OpenRecordset(query, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);
                
                // Check for existence of suite id in the database. Add if new
                if (rs.EOF)
                {
                    AddRecord(rs, recordsExcel);
                    recordsExcel.MoveNext();
                    rs.Close();
                    continue;
                }

                rs.Edit();

                // We don't update Suite ID, Investigator, Notes, Status, or Activity
                // Investigator, Notes, Status, Activity do not come from results data.
                rs.Fields["Suite Name"].Value = recordsExcel.Fields["Name"].Value;
                rs.Fields["Defect"].Value = recordsExcel.Fields["Defect"].Value;
                rs.Fields["Machine"].Value = recordsExcel.Fields["Machine"].Value;
                rs.Fields["Test Time"].Value = recordsExcel.Fields["Test Time"].Value;
                rs.Fields["Result"].Value = recordsExcel.Fields["Result"].Value;
                rs.Fields["Org"].Value = recordsExcel.Fields["Org"].Value;
                rs.Fields["Platform"].Value = recordsExcel.Fields["Platform"].Value;
                rs.Fields["Simulation"].Value = recordsExcel.Fields["Simulation"].Value;
                rs.Fields["User"].Value = recordsExcel.Fields["User"].Value;
                rs.Fields["Kit Type"].Value = recordsExcel.Fields["Kit Type"].Value;
                rs.Fields["OS"].Value = recordsExcel.Fields["OS"].Value;
                rs.Fields["Office"].Value = recordsExcel.Fields["Office"].Value;
                rs.Fields["Kit"].Value = recordsExcel.Fields["Kit"].Value;
                rs.Fields["First Fail"].Value = recordsExcel.Fields["First Fail"].Value;

                rs.Update(1,false);

                recordsExcel.MoveNext();

                rs.Close();
            }
        }

        private void UpdateFromParentFile(Database db, string accessTable, string xlsxFile)
        {
            // get count of tests
            int totalcount = TestCount(xlsxFile);
            int currentcount = 0;

            // really need some sort of using block around the excel access...
            Recordset recordsExcel = OpenExcelRecords(xlsxFile);

            Console.WriteLine($"\n\nUpdating database with {xlsxFile}");

            while (!recordsExcel.EOF)
            {
                currentcount++;
                Console.Write($"\r Updating {currentcount} of {totalcount}");

                object suiteID = recordsExcel.Fields[SUITEID].Value; // suiteid comes from excel as a double

                // query for the suite id record in the access database
                string query = $"SELECT * FROM {accessTable} WHERE [Suite ID] = {suiteID}";
                Recordset rs = db.OpenRecordset(query, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);

                // If we find a test in the parent data that is not in the database then we are probably 
                // using the wrong date for the parent data since the rebase should have brought over new tests.
                if (rs.EOF)
                {
                    errors.Add($"Suite ID [{suiteID}] not found in database. Possible use of incorrect parent branch result data");
                    recordsExcel.MoveNext();
                    rs.Close();
                    continue;
                }

                rs.Edit();

                // When updating parent data we only need result and kit
                rs.Fields["Parent Result"].Value = recordsExcel.Fields["Result"].Value;
                rs.Fields["Parent Kit"].Value = recordsExcel.Fields["Kit"].Value;

                rs.Update(1, false);

                recordsExcel.MoveNext();
                rs.Close();
            }
        }

        private void AddRecord(Recordset recordsDb, Recordset recordsExcel)
        {
            recordsDb.AddNew();

            // Suite ID
            // Suite Name
            // Defect
            // Investigator
            // Test Time
            // Result
            // Parent Result
            // Org
            // Platform
            // Simulation
            // User
            // Kit Type
            // OS
            // Office
            // Kit
            // Parent Kit
            // First Fail
            // Notes
            // Activity
            // Status

            recordsDb.Fields["Suite ID"].Value = recordsExcel.Fields["Suite ID"].Value;
            recordsDb.Fields["Suite Name"].Value = recordsExcel.Fields["Name"].Value;
            recordsDb.Fields["Defect"].Value = recordsExcel.Fields["Defect"].Value;
            recordsDb.Fields["Investigator"].Value = "unassigned";
            recordsDb.Fields["Machine"].Value = recordsExcel.Fields["Machine"].Value;
            recordsDb.Fields["Test Time"].Value = recordsExcel.Fields["Test Time"].Value;
            recordsDb.Fields["Result"].Value = recordsExcel.Fields["Result"].Value;
            recordsDb.Fields["Parent Result"].Value = "";  // For new records set parent result empty
            recordsDb.Fields["Org"].Value = recordsExcel.Fields["Org"].Value;
            recordsDb.Fields["Platform"].Value = recordsExcel.Fields["Platform"].Value;
            recordsDb.Fields["Simulation"].Value = recordsExcel.Fields["Simulation"].Value;
            recordsDb.Fields["User"].Value = recordsExcel.Fields["User"].Value;
            recordsDb.Fields["Kit Type"].Value = recordsExcel.Fields["Kit Type"].Value;
            recordsDb.Fields["OS"].Value = recordsExcel.Fields["OS"].Value;
            recordsDb.Fields["Office"].Value = recordsExcel.Fields["Office"].Value;
            recordsDb.Fields["Kit"].Value = recordsExcel.Fields["Kit"].Value;
            recordsDb.Fields["Parent Kit"].Value = ""; // For new records set parent kit empty
            recordsDb.Fields["First Fail"].Value = recordsExcel.Fields["First Fail"].Value;
            recordsDb.Fields["Notes"].Value = "";
            recordsDb.Fields["Activity"].Value = "";
            recordsDb.Fields["Status"].Value = "";

            recordsDb.Update(1,false);
        }

        public Recordset OpenExcelRecords(string xlsxFile)
        {
            Database dbExcel;
            DBEngine dben = new DBEngineClass();

            dbExcel = dben.OpenDatabase(xlsxFile, false, true, "Excel 12.0 Xml;HDR=YES;");

            int c = dbExcel.TableDefs.Count;

            // Error if count is greater than 1...
            if (c != 1)
            {
                dbExcel.Close();
                throw new Exceptions.ExcelSheetCountException(xlsxFile);
            }

            string sheetName = dbExcel.TableDefs[0].Name;

            Recordset rs = dbExcel.OpenRecordset(sheetName, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);

            return rs;
        }

        public int TestCount(string xlsxFile)
        {
            Database dbExcel;
            DBEngine dben = new DBEngineClass();

            dbExcel = dben.OpenDatabase(xlsxFile, false, true, "Excel 12.0 Xml;HDR=YES;");

            int c = dbExcel.TableDefs.Count;
            // Error if count is greater than 1...
            if (c != 1)
            {
                dbExcel.Close();
                throw new Exceptions.ExcelSheetCountException(xlsxFile);
            }

            string sheetName = dbExcel.TableDefs[0].Name;

            string query = $"SELECT Count(*) as [CountOfRows] FROM [{sheetName}]";
            Recordset recordsCount = dbExcel.OpenRecordset(query, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);
            int count = (int)recordsCount.Fields["CountOfRows"].Value;

            recordsCount.Close();
            dbExcel.Close();

            if (count <= 0)
            {
                throw new Exceptions.ExcelTestCountException(xlsxFile);
            }

            return count;
        }
    }
}
