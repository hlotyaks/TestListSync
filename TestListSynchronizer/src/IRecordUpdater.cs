namespace TestListSynchronizer
{
    public interface IRecordUpdater
    {
        void AddIncommingRecord(int suiteID);
        void NewRecord(IRecords dbRecords, IRecords xlRecords);
        void NewRecord(IRecords dbRecords, SuiteResult testResult);
        void NotRunRecord(int suiteID, IRecords dbRecords);
        object SetValueOrDefault(object value);
        void UpdateParentRecord(IRecords dbRecords, IRecords xlRecords);
        void UpdateParentRecord(IRecords dbRecords, SuiteResult testResult);

        void UpdateRecord(IRecords dbRecords, IRecords xlRecords);
        void UpdateRecord(IRecords dbRecords, SuiteResult testResult);

    }
}