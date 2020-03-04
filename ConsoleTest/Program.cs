using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {

            string table = "TestList";
            string db = @"C:\Users\hlotyaks\Source\Repos\TestListSync\Test\10-20-73-TestList.accdb";
            string asrtxlsx = @"C:\Users\hlotyaks\Source\Repos\TestListSync\Test\mar03asrt.xlsx";
            string bfrxlsx = @"C:\Users\hlotyaks\Source\Repos\TestListSync\Test\mar03bfr.xlsx";


            TestListSynchronizer.DatabaseSync dbsync = new TestListSynchronizer.DatabaseSync(db, table);
            dbsync.UpdateDatabase(asrtxlsx, bfrxlsx);

        }
    }
}
