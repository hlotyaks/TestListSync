namespace TestListSynchronizer
{
    public interface IRecordUpdater
    {
        void AddIncommingRecord(int suiteID);
        void NewRecord(IRecords dbRecords, IRecords xlRecords);
        void NotRunRecord(int suiteID, IRecords dbRecords);
        object SetValueOrDefault(object value);
        void UpdateParentRecord(IRecords dbRecords, IRecords xlRecords);
        void UpdateRecord(IRecords dbRecords, IRecords xlRecords);
    }
}