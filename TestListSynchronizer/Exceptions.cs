using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestListSynchronizer.Exceptions
{
    public class ExcelSheetCountException : Exception
    {
        public ExcelSheetCountException(string message) : base(message)
        { }
    }

    public class ExcelTestCountException : Exception
    {
        public ExcelTestCountException(string message) : base(message)
        { }
    }
}
