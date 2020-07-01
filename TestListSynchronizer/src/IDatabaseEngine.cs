namespace TestListSynchronizer
{
    public interface IDatabaseEngine
    {
        IDatabase Open(string name);
        IDatabase Open(string name, string connectionString);
    }
}