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
            string db = @"C:\Users\hlotyaks\Source\Repos\TestListUpdater\Test\10-20-73-TestList.accdb";
            string asrtxlsx = @"C:\Users\hlotyaks\Source\Repos\TestListUpdater\Test\feb26asrt.xlsx";
            string bfrxlsx = @"C:\Users\hlotyaks\Source\Repos\TestListUpdater\Test\feb26bfr.xlsx";


            TestListSynchronizer.DataBaseConnection dbc = new TestListSynchronizer.DataBaseConnection(db, table, asrtxlsx, bfrxlsx);
        }
    }
}
