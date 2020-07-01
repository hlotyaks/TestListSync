using System;

namespace TestListSynchronizer
{
    public interface IRecords : IDisposable
    {
        bool EOF { get; }

        void Add();
        void Edit();
        object GetFieldValue(string fieldName);
        void MoveNext();
        void SetFieldValue(string fieldName, object value);
        void Update();
        int GetSuiteID();
    }
}