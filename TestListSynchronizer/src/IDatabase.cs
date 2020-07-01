using System;

namespace TestListSynchronizer
{
    public interface IDatabase : IDisposable
    {
        int TableCount { get; }

        void Open(string name);
        void Open(string name, string connectionString);
        IRecords OpenRecords(string queryString);
        string TableName(int index);
        int TableSize(string table);
    }
}