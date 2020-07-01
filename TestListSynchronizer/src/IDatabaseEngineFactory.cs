namespace TestListSynchronizer
{
    public interface IDatabaseEngineFactory
    {
        IDatabaseEngine CreateDatabaseEngine();
        IRecordUpdater CreateRecordUpdater();
    }
}