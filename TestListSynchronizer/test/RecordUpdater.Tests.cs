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
    public class RecordUpdaterTests
    {
        IRecordUpdater _recordUpdater;

        [TestInitialize]
        public void TestSetup()
        {
            _recordUpdater = new RecordUpdater() as IRecordUpdater;
        }

        [TestMethod]
        public void TestSetValueOrDefault1()
        {
            object value = _recordUpdater.SetValueOrDefault(1);

            value.Should().Be(1);

        }

        [TestMethod]
        public void TestSetValueOrDefault2()
        {
            object value = _recordUpdater.SetValueOrDefault(String.Empty);

            value.Should().Be(" ");
        }

        [TestMethod]
        public void TestSetValueOrDefault3()
        {
            string s = null;

            object value = _recordUpdater.SetValueOrDefault(s);

            value.Should().Be(" ");
        }

        [TestMethod]
        public void TestSetValueOrDefault4()
        {
            object value = _recordUpdater.SetValueOrDefault("test");

            value.Should().Be("test");
        }

        [TestMethod]
        public void TestNotRunRecord1()
        {
            Mock<IRecords> mockRecords = new Mock<IRecords>();
            _recordUpdater.AddIncommingRecord(1);

            _recordUpdater.NotRunRecord(1, mockRecords.Object);

            mockRecords.Verify(x => x.Edit(), Times.Never());
            mockRecords.Verify(x => x.Update(), Times.Never());
        }

        [TestMethod]
        public void TestNotRunRecord2()
        {
            Mock<IRecords> mockRecords = new Mock<IRecords>();
            _recordUpdater.AddIncommingRecord(1);

            _recordUpdater.NotRunRecord(2, mockRecords.Object);

            mockRecords.Verify(x => x.Edit());
            mockRecords.Verify(x => x.Update());
        }

        [TestMethod]
        public void TestNewRecord()
        {
            Mock<IRecords> mockRecords = new Mock<IRecords>();

            _recordUpdater.NewRecord(mockRecords.Object, mockRecords.Object);

            mockRecords.Verify(x => x.Add());
            mockRecords.Verify(x => x.Update());
        }

        [TestMethod]
        public void TestUpdateRecord()
        {
            Mock<IRecords> mockRecords = new Mock<IRecords>();

            _recordUpdater.UpdateRecord(mockRecords.Object, mockRecords.Object);

            mockRecords.Verify(x => x.Edit());
            mockRecords.Verify(x => x.Update());
        }

        [TestMethod]
        public void TestUpdateParentRecord()
        {
            Mock<IRecords> mockRecords = new Mock<IRecords>();

            _recordUpdater.UpdateParentRecord(mockRecords.Object, mockRecords.Object);

            mockRecords.Verify(x => x.Edit());
            mockRecords.Verify(x => x.Update());
        }
    }
}
