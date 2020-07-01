using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace TestListSynchronizer
{
    public class DatabaseEngine : IDatabaseEngine
    {
        DAO.DBEngineClass engine;

        /// <summary>
        /// 
        /// </summary>
        public DatabaseEngine()
        {
            // requires Microsoft.Office.Interop.Access.Dao\14.0.0.0__71e9bce111e9429c
            engine = new DAO.DBEngineClass();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        public IDatabase Open(string name)
        {
            Database db = new Database(engine);
            db.Open(name);
            return db;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        public IDatabase Open(string name, string connectionString)
        {
            Database db = new Database(engine);
            db.Open(name, connectionString);
            return db;
        }

    }
}
