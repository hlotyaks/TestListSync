using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TestListSynchronizer;
using Moq;
using FluentAssertions;
using System.Text.RegularExpressions;

namespace UnitTest
{
    [TestClass]
    public class DatabaseSyncTests
    {
        DatabaseSync dbsync;
        Mock<IDatabaseEngineFactory> mockFactory;
        Mock<IDatabaseEngine> mockEngine;
        Mock<IDatabase> mockXLDB;
        Mock<IDatabase> mockDB;
        Mock<IRecords> mockXLRecords;
        Mock<IRecords> mockDBRecords;
        Mock<IRecordUpdater> mockRecordUpdater;
        string database = "db";
        string dbtable = "dbtable";

        [TestInitialize]
        public void TestSetup()
        {
            mockFactory = new Mock<IDatabaseEngineFactory>();
            mockEngine = new Mock<IDatabaseEngine>();
            mockXLDB = new Mock<IDatabase>();
            mockDB = new Mock<IDatabase>();
            mockXLRecords = new Mock<IRecords>();
            mockDBRecords = new Mock<IRecords>();
            mockRecordUpdater = new Mock<IRecordUpdater>();

            mockFactory.Setup(x => x.CreateDatabaseEngine()).Returns(mockEngine.Object);
            mockFactory.Setup(x => x.CreateRecordUpdater()).Returns(mockRecordUpdater.Object);

            dbsync = new DatabaseSync(database, dbtable, mockFactory.Object);
            dbsync.Should().NotBeNull();
        }

        [TestMethod]
        public void UpdateDatabase1()
        {
            List<string> filelist = new List<string>(new[] { "file1" });
            List<int> incommingSuiteIDs = new List<int>(new[] { 1, 2 });
            List<int> dbSuiteIDs = new List<int>(new[] { 1 });

            int currentSuiteID=0;
            string fileName = "file1";
            string tableName = "table";
            int tableCount = 1;
            int testCount = 1;
            int xlidx = 0;
            int dbidx = 0;

            mockEngine.Setup(x => x.Open(database)).Returns(mockDB.Object);

            MockTestCountCall(fileName, tableName, tableCount, testCount);
            MockUpdateDisabledTests(incommingSuiteIDs);

            mockEngine.Setup(x => x.Open(fileName, "Excel 12.0 Xml;HDR=YES;")).Returns(mockXLDB.Object);
            mockXLDB.Setup(x => x.TableCount).Returns(tableCount);
            mockXLDB.Setup(x => x.OpenRecords("table")).Returns(mockXLRecords.Object);

            mockXLRecords.Setup(x => x.EOF).Returns(() =>
                {
                    if (xlidx < incommingSuiteIDs.Count)
                    {
                        return false;
                    }

                    return true;
                });

            mockXLRecords.Setup(x => x.GetSuiteID()).Returns(() =>
                {
                    currentSuiteID = incommingSuiteIDs[xlidx++];
                    return currentSuiteID;
                });

            mockRecordUpdater.Setup(x => x.AddIncommingRecord(It.IsAny<int>()));

            mockDB.Setup(x => x.OpenRecords(It.IsRegex(@"SELECT \* FROM dbtable WHERE \[Suite ID] = \d*"))).Returns(mockDBRecords.Object);

            mockDBRecords.Setup(x => x.EOF).Returns(() =>
            {
                if (dbSuiteIDs.Contains(currentSuiteID))
                {
                    return false;
                }

                return true;
            });

            dbsync.UpdateDatabase(filelist);

            mockRecordUpdater.Verify(x => x.NewRecord(It.IsAny<IRecords>(), It.IsAny<IRecords>()));
            mockRecordUpdater.Verify(x => x.UpdateRecord(It.IsAny<IRecords>(), It.IsAny<IRecords>()));


            //List<string> filelist = new List<string>();
            //filelist.Add("file1");
            //List<int> suiteIDs = new List<int>(new[] { 1, 2 });
            //int xlidx = 0;
            //int dbidx = 0;


            ////int suiteID = 1;
            ////string suiteQuery = $"SELECT * FROM {dbtable} WHERE [Suite ID] = {suiteID}";
            //bool firstTime = true;

            //mockEngine.Setup(x => x.Open(database)).Returns(mockDB.Object);

            //MockTestCountCall(fileName, tableName, tableCount, testCount);
            ////MockUpdateDisabledTests(new Queue<int>(new[] { 1, 2 }));

            //mockEngine.Setup(x => x.Open(fileName, "Excel 12.0 Xml;HDR=YES;")).Returns(mockXLDB.Object);
            //mockXLDB.Setup(x => x.TableCount).Returns(tableCount);
            //mockXLDB.Setup(x => x.OpenRecords("table")).Returns(mockXLRecords.Object);
            //mockXLRecords.Setup(x => x.EOF).Returns(() =>
            //    {
            //        if (xlidx < suiteIDs.Count)
            //        {
            //            return false;
            //        }

            //        return true;
            //    });

            //mockXLRecords.Setup(x => x.GetFieldValue("Suite ID")).Returns(() =>
            //    {
            //        double suiteID = (double)suiteIDs[xlidx++];
            //        return suiteID;
            //    });


            //mockDB.Setup(x => x.OpenRecords("SELECT * FROM dbtable")).Returns(mockDBRecords.Object);

            ////SELECT * FROM dbtable WHERE [Suite ID] = 1//
            //Regex rx = new Regex(@"SELECT \* FROM dbtable WHERE [Suite ID] = \d*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            //Regex rx2 = new Regex(@"SELECT \* FROM dbtable WHERE \[Suite ID] = \d*");

            //string text = "SELECT * FROM dbtable WHERE [Suite ID] = 1";
            //string text2 = "SELECT * FROM dbtable WHERE [Suite ID] = 1";
            //MatchCollection matches = rx.Matches(text);
            //matches = rx2.Matches(text2);

            //mockDB.Setup(x => x.OpenRecords(It.IsRegex(@"SELECT \* FROM dbtable WHERE \[Suite ID] = \d*"))).Returns(mockDBRecords.Object);
            //mockDB.Setup(x => x.OpenRecords(It.IsRegex(@"SELECT \* FROM dbtable"))).Returns(mockDBRecords.Object);

            //mockDBRecords.Setup(x => x.EOF).Returns(() =>
            //{
            //    if (dbidx < suiteIDs.Count)
            //    {
            //        return false;
            //    }

            //    return true;
            //});

            //mockDBRecords.Setup(x => x.GetFieldValue("Suite ID")).Returns(() =>
            //{
            //    double suiteID = (double)suiteIDs[dbidx++];
            //    return suiteID;
            //});

            //mockDBRecords.Setup(x => x.EOF).Returns(false);      

            //dbsync.UpdateDatabase(filelist);

            //mockXLRecords.Verify(x => x.MoveNext());
            //mockDBRecords.Verify(x => x.Edit());
            //mockDBRecords.Verify(x => x.Update());
        }

        [TestMethod]
        public void TestCount1()
        {
            string fileName = "file1";
            string tableName = "table";
            int tableCount = 1;
            int testCount = 1;

            MockTestCountCall(fileName, tableName, tableCount, testCount);

            int value = dbsync.TestCount(fileName);
            value.Should().Be(testCount);

        }

        [TestMethod]
        [ExpectedException(typeof(TestListSynchronizer.Exceptions.ExcelSheetCountException))]
        public void TestCount2()
        {
            string fileName = "file1";
            string tableName = "table";
            int tableCount = 0;
            int testCount = 1;

            MockTestCountCall(fileName, tableName, tableCount, testCount);

            int value = dbsync.TestCount(fileName);
            value.Should().Be(testCount);

        }
        [TestMethod]
        [ExpectedException(typeof(TestListSynchronizer.Exceptions.ExcelTestCountException))]
        public void TestCount3()
        {
            string fileName = "file1";
            string tableName = "table";
            int tableCount = 1;
            int testCount = -1;

            MockTestCountCall(fileName, tableName, tableCount, testCount);

            int value = dbsync.TestCount(fileName);
            value.Should().Be(testCount);

        }

        private void MockTestCountCall(string filename, string tablename, int tablecount, int testcount)
        {
            string countQuery = $"SELECT Count(*) as [CountOfRows] FROM [{tablename}]";

            mockEngine.Setup(x => x.Open(filename, "Excel 12.0 Xml;HDR=YES;")).Returns(mockXLDB.Object);
            mockXLDB.Setup(x => x.TableCount).Returns(tablecount);
            mockXLDB.Setup(x => x.TableName(0)).Returns(tablename);
            mockXLDB.Setup(x => x.OpenRecords(countQuery)).Returns(mockXLRecords.Object);
            mockXLRecords.Setup(x => x.GetFieldValue("CountOfRows")).Returns(testcount);
        }

        private void MockUpdateDisabledTests(List<int> suiteIDs)
        {
            Mock<IRecords> mockDBRecords2 = new Mock<IRecords>();
            string query = $"SELECT * FROM {dbtable}";
            int idx = 0;
            int currentSuiteID = 0;

            mockDB.Setup(x => x.TableSize(It.IsAny<string>())).Returns(1);
            mockDB.Setup(x => x.OpenRecords(query)).Returns(mockDBRecords2.Object);
            mockDBRecords2.Setup(x => x.EOF).Returns(() =>
            {
                if (idx < suiteIDs.Count)
                {
                    return false;
                }

                return true;
            });

            mockDBRecords2.Setup(x => x.GetSuiteID()).Returns(() =>
            {
                currentSuiteID = suiteIDs[idx++];
                return currentSuiteID;
            });
        }
    }
}
