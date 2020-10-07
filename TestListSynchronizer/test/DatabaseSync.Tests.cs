using System;
using System.Linq;
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
        Mock<ITestListSyncFactory> mockFactory;
        Mock<IDatabaseEngine> mockEngine;
        Mock<IDatabase> mockXLDB;
        Mock<IDatabase> mockDB;
        Mock<IRecords> mockXLRecords;
        Mock<IRecords> mockDBRecords;
        Mock<IRecordUpdater> mockRecordUpdater;
        Mock<IJarvisWrapper> mockJarvis;
        string database = "db";
        string dbtable = "dbtable";
        string jarvisApp = "jarvis.exe";

        [TestInitialize]
        public void TestSetup()
        {
            mockFactory = new Mock<ITestListSyncFactory>();
            mockEngine = new Mock<IDatabaseEngine>();
            mockXLDB = new Mock<IDatabase>();
            mockDB = new Mock<IDatabase>();
            mockXLRecords = new Mock<IRecords>();
            mockDBRecords = new Mock<IRecords>();
            mockRecordUpdater = new Mock<IRecordUpdater>();
            mockJarvis = new Mock<IJarvisWrapper>();

            mockFactory.Setup(x => x.CreateDatabaseEngine()).Returns(mockEngine.Object);
            mockFactory.Setup(x => x.CreateRecordUpdater()).Returns(mockRecordUpdater.Object);
            mockFactory.Setup(x => x.CreateJarvisWrapper(jarvisApp)).Returns(mockJarvis.Object);

            dbsync = new DatabaseSync(database, dbtable, mockFactory.Object, jarvisApp);
            dbsync.Should().NotBeNull();
        }

        [TestMethod]
        public void UpdateDatabaseFromProjectOneSuite()
        {
            string project = "test-project";
            string baseline = null;
            string parentProject = null;
            string parentBaseline = null;
            int suiteIDX = 0;
            List<SuiteResult> testdata = new List<SuiteResult>();

            testdata.Add(CreateTestData(1));
            List<int> incommingSuiteIDs = testdata.Select(i => i.SuiteID).ToList();
            List<int> dbSuiteIDs = new List<int>(new[] { 1 });

            MockUpdateDisabledTests(incommingSuiteIDs);

            mockEngine.Setup(x => x.Open(database)).Returns(mockDB.Object);
            mockJarvis.Setup(x => x.FetchResults(project)).Returns(testdata);

            mockDB.Setup(x => x.OpenRecords(It.IsRegex(@"SELECT \* FROM dbtable WHERE \[Suite ID] = \d*"))).Returns(mockDBRecords.Object);

            mockDBRecords.Setup(x => x.EOF).Returns(() =>
            {
                if (dbSuiteIDs.Contains(incommingSuiteIDs[suiteIDX]))
                {
                    suiteIDX++;
                    return false;
                }

                return true;
            });


            dbsync.UpdateDatabase(project, baseline, parentProject, parentBaseline);

            mockRecordUpdater.Verify(x => x.NewRecord(It.IsAny<IRecords>(), It.IsAny<SuiteResult>()), Times.Exactly(0));
            mockRecordUpdater.Verify(x => x.UpdateRecord(It.IsAny<IRecords>(), It.IsAny<SuiteResult>()), Times.Exactly(1));

        }

        [TestMethod]
        public void UpdateDatabaseFromProjectTwoSuites()
        {
            string project = "test-project";
            string baseline = null;
            string parentProject = null;
            string parentBaseline = null;
            int suiteIDX = 0;
            List<SuiteResult> testdata = new List<SuiteResult>();

            testdata.Add(CreateTestData(1));
            testdata.Add(CreateTestData(2));
            List<int> incommingSuiteIDs = testdata.Select(i => i.SuiteID).ToList();
            List<int> dbSuiteIDs = new List<int>(new[] { 1 });

            MockUpdateDisabledTests(incommingSuiteIDs);

            mockEngine.Setup(x => x.Open(database)).Returns(mockDB.Object);
            mockJarvis.Setup(x => x.FetchResults(project)).Returns(testdata);

            mockDB.Setup(x => x.OpenRecords(It.IsRegex(@"SELECT \* FROM dbtable WHERE \[Suite ID] = \d*"))).Returns(mockDBRecords.Object);

            mockDBRecords.Setup(x => x.EOF).Returns(() =>
            {
                if (dbSuiteIDs.Contains(incommingSuiteIDs[suiteIDX]))
                {
                    suiteIDX++;
                    return false;
                }

                return true;
            });


            dbsync.UpdateDatabase(project, baseline, parentProject, parentBaseline);

            mockRecordUpdater.Verify(x => x.NewRecord(It.IsAny<IRecords>(), It.IsAny<SuiteResult>()), Times.Exactly(1));
            mockRecordUpdater.Verify(x => x.UpdateRecord(It.IsAny<IRecords>(), It.IsAny<SuiteResult>()), Times.Exactly(1));

        }

        [TestMethod]
        public void UpdateDatabaseFromProjectTwoSuitesAndParentData()
        {
            string project = "test-project";
            string baseline = null;
            string parentProject = "parent-project";
            string parentBaseline = "parent-baseline";
            int suiteIDX = 0;
            List<SuiteResult> testdata = new List<SuiteResult>();

            testdata.Add(CreateTestData(1));
            testdata.Add(CreateTestData(2));
            List<int> incommingSuiteIDs = testdata.Select(i => i.SuiteID).ToList();
            List<int> dbSuiteIDs = testdata.Select(i => i.SuiteID).ToList();

            MockUpdateDisabledTests(incommingSuiteIDs);

            mockEngine.Setup(x => x.Open(database)).Returns(mockDB.Object);
            mockJarvis.Setup(x => x.FetchResults(project)).Returns(testdata);
            mockJarvis.Setup(x => x.FetchResults(parentProject, parentBaseline)).Returns(testdata);

            mockDB.Setup(x => x.OpenRecords(It.IsRegex(@"SELECT \* FROM dbtable WHERE \[Suite ID] = \d*"))).Returns(mockDBRecords.Object);

            mockDBRecords.Setup(x => x.EOF).Returns(() =>
            {
                if (dbSuiteIDs.Contains(incommingSuiteIDs[suiteIDX]))
                {
                    suiteIDX++;
                    if (suiteIDX == testdata.Count)
                        suiteIDX = 0;
                    return false;
                }

                return true;
            });


            dbsync.UpdateDatabase(project, baseline, parentProject, parentBaseline);

            mockRecordUpdater.Verify(x => x.NewRecord(It.IsAny<IRecords>(), It.IsAny<SuiteResult>()), Times.Exactly(0));
            mockRecordUpdater.Verify(x => x.UpdateRecord(It.IsAny<IRecords>(), It.IsAny<SuiteResult>()), Times.Exactly(2));
            mockRecordUpdater.Verify(x => x.UpdateParentRecord(It.IsAny<IRecords>(), It.IsAny<SuiteResult>()), Times.Exactly(2));

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
        private SuiteResult CreateTestData(int suiteid)
        {

            SuiteResult s1 = new SuiteResult();
            s1.Defect = "defect";
            s1.ElapsedTime = 1000;
            s1.EnvironmentTags = new string[] { "ASRT", "Windows10", "64bit", "Office2016" };
            s1.FirstFail = "Nov3-20";
            s1.KitDate = "Nov3-20";
            s1.MachineName = "bob";
            s1.Organization = "QA";
            s1.Platform = "UF";
            s1.Result = "Passed";
            s1.SimulationType = "none";
            s1.SuiteID = suiteid;
            s1.SuiteName = "suite name";
            s1.User = "hal";

            return s1;

        }
    }
   

}
