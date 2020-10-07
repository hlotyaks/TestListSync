namespace TestListSynchronizer
{
    public interface ITestListSyncFactory
    {
        IDatabaseEngine CreateDatabaseEngine();
        IRecordUpdater CreateRecordUpdater();
        IJarvisWrapper CreateJarvisWrapper(string jarvisApp);

    }
}