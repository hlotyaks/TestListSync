using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestListSynchronizer
{
    public class TestListSyncFactory : ITestListSyncFactory
    {
        public IDatabaseEngine CreateDatabaseEngine()
        {
            return new DatabaseEngine() as IDatabaseEngine;
        }

        public IJarvisWrapper CreateJarvisWrapper(string jarvisApp)
        {
            return new JarvisWrapper(jarvisApp) as IJarvisWrapper;
        }

        public IRecordUpdater CreateRecordUpdater()
        {
            return new RecordUpdater() as IRecordUpdater;
        }
    }
}
