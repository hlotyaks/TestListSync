using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestListSynchronizer
{
    public interface IJarvisWrapper
    {
        List<SuiteResult> FetchResults(string project);

        List<SuiteResult> FetchResults(string project, string baseline);
    }
}
