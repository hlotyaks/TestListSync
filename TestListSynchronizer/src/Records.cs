using System;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace TestListSynchronizer
{
    public class Records : IRecords
    {
        private const string SUITEID = "Suite ID";
        DAO.Recordset records;
        bool disposed = false;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="database"></param>
        /// <param name="queryString"></param>
        public Records(DAO.Database database, string queryString)
        {
            records = database.OpenRecordset(queryString, DAO.RecordsetTypeEnum.dbOpenDynaset, null, DAO.LockTypeEnum.dbOptimistic);
        }

        /// <summary>
        /// Specialized GetFieldValue for SuiteID
        /// </summary>
        /// <returns>suiteID</returns>
        public int GetSuiteID()
        {
            return (int)((double)GetFieldValue(SUITEID));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public object GetFieldValue(string fieldName)
        {
            return records.Fields[fieldName].Value;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fieldName"></param>
        /// <param name="value"></param>
        public void SetFieldValue(string fieldName, object value)
        {
            records.Fields[fieldName].Value = value;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Add()
        {
            records.AddNew();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Edit()
        {
            records.Edit();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Update()
        {
            records.Update(1, false);
        }

        /// <summary>
        /// 
        /// </summary>
        public void MoveNext()
        {
            records.MoveNext();
        }

        /// <summary>
        /// 
        /// </summary>
        public bool EOF
        {
            get
            {
                return records.EOF;
            }
        }


        #region Dispose support

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            Dispose(true);

            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 
        /// </summary>
        ~Records()
        {
            Dispose(false);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    // Close the record set.
                    records.Close();

                }

                // Note disposing has been done.
                disposed = true;
            }
        }
        #endregion
    }
}