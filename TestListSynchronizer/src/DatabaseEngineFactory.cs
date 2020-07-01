using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestListSynchronizer
{
    public class DatabaseEngineFactory : IDatabaseEngineFactory
    {
        public IDatabaseEngine CreateDatabaseEngine()
        {
            return new DatabaseEngine() as IDatabaseEngine;
        }

        public IRecordUpdater CreateRecordUpdater()
        {
            return new RecordUpdater() as IRecordUpdater;
        }
    }
}
