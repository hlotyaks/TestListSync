using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace TestListSynchronizer
{
    public class Database : IDisposable, IDatabase
    {
        DAO.DBEngineClass engine;
        DAO.Database access;
        bool disposed = false;

        public Database(DAO.DBEngineClass engine)
        {
            this.engine = engine;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        public void Open(string name)
        {
            try
            {
                access = engine.OpenDatabase(name);
            }
            catch (Exception)
            {
                throw new Exceptions.DatabaseOpenException(name);
            }

            // Refresh the data. will pull from sharepoint.
            access.TableDefs.Refresh();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="connectionString"></param>
        public void Open(string name, string connectionString)
        {
            try
            {
                access = engine.OpenDatabase(name, false, true, connectionString);
            }
            catch (Exception)
            {
                throw new Exceptions.DatabaseOpenException(name);
            }

            // Refresh the data. will pull from sharepoint.
            access.TableDefs.Refresh();
        }

        /// <summary>
        /// 
        /// </summary>
        public int TableCount
        {
            get => access.TableDefs.Count;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public string TableName(int index)
        {
            return access.TableDefs[index].Name;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public int TableSize(string table)
        {
            string query = $"SELECT Count(*) as [CountOfRows] FROM {table}";
            DAO.Recordset recordsCount = access.OpenRecordset(query, DAO.RecordsetTypeEnum.dbOpenDynaset, null, DAO.LockTypeEnum.dbOptimistic);
            int count = (int)recordsCount.Fields["CountOfRows"].Value;
            recordsCount.Close();
            return count;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="queryString"></param>
        /// <returns></returns>
        public IRecords OpenRecords(string queryString)
        {
            return new Records(access, queryString);
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
                    // Dispose managed resources.
                    // Nothing to do atm

                }

                // Close the database.
                access.Close();

                // Note disposing has been done.
                disposed = true;
            }
        }

        #endregion

    }
}