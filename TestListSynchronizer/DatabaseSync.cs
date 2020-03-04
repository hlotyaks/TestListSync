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

        public void UpdateDatabase(string asrtxlsx, string bfrxlsx)
        {
            dbAcess = dbEn.OpenDatabase(dbName);

            // Refresh the data. will pull from sharepoint.
            dbAcess.TableDefs.Refresh();

            UpdateFromExcel(dbAcess, dbTable, asrtxlsx);
            UpdateFromExcel(dbAcess, dbTable, bfrxlsx);
            UpdateDisabledTests(dbAcess, dbTable);

            dbAcess.Close();
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

        private void UpdateFromExcel(Database db, string accessTable, string  excelFile)
        {
            // get count of tests
            int totalcount = TestCount(excelFile);
            int currentcount = 0;
            // really need some sort oif using block around the excel access...
            Recordset recordsExcel = OpenExcelRecords(excelFile);

            Console.WriteLine($"\n\nUpdating database with {excelFile}");

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

        private void AddRecord(Recordset recordsDb, Recordset recordsExcel)
        {
            recordsDb.AddNew();

            // Suite ID
            // Suite Name
            // Defect
            // Investigator
            // Test Time
            // Result
            // Org
            // Platform
            // Simulation
            // User
            // Kit Type
            // OS
            // Office
            // Kit
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
                recordsDb.Fields["Org"].Value = recordsExcel.Fields["Org"].Value;
                recordsDb.Fields["Platform"].Value = recordsExcel.Fields["Platform"].Value;
                recordsDb.Fields["Simulation"].Value = recordsExcel.Fields["Simulation"].Value;
                recordsDb.Fields["User"].Value = recordsExcel.Fields["User"].Value;
                recordsDb.Fields["Kit Type"].Value = recordsExcel.Fields["Kit Type"].Value;
                recordsDb.Fields["OS"].Value = recordsExcel.Fields["OS"].Value;
                recordsDb.Fields["Office"].Value = recordsExcel.Fields["Office"].Value;
                recordsDb.Fields["Kit"].Value = recordsExcel.Fields["Kit"].Value;
                recordsDb.Fields["First Fail"].Value = recordsExcel.Fields["First Fail"].Value;
                recordsDb.Fields["Notes"].Value = "";
                recordsDb.Fields["Activity"].Value = "";
                recordsDb.Fields["Status"].Value = "";

            recordsDb.Update(1,false);
        }

        public Recordset OpenExcelRecords(string file)
        {
            Database dbExcel;
            DBEngine dben = new DBEngineClass();

            dbExcel = dben.OpenDatabase(file, false, true, "Excel 12.0 Xml;HDR=YES;");

            int c = dbExcel.TableDefs.Count;

            // Error if count is greater than 1...

            string sheetName = dbExcel.TableDefs[0].Name;

            Recordset rs = dbExcel.OpenRecordset(sheetName, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);

            return rs;
        }

        public int TestCount(string excelFile)
        {
            Database dbExcel;
            DBEngine dben = new DBEngineClass();

            dbExcel = dben.OpenDatabase(excelFile, false, true, "Excel 12.0 Xml;HDR=YES;");

            int c = dbExcel.TableDefs.Count;
            // Error if count is greater than 1...
            string sheetName = dbExcel.TableDefs[0].Name;

            string query = $"SELECT Count(*) as [CountOfRows] FROM [{sheetName}]";
            Recordset recordsCount = dbExcel.OpenRecordset(query, RecordsetTypeEnum.dbOpenDynaset, null, LockTypeEnum.dbOptimistic);
            int count = (int)recordsCount.Fields["CountOfRows"].Value;

            recordsCount.Close();
            dbExcel.Close();

            return count;
        }
    }
}
