using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestListSynchronizer
{
    public class SuiteResult
    {
        public int SuiteID { get; set; }
        public string Result { get; set; }
        public string Organization { get; set; }
        public string SuiteName { get; set; }
        public string Defect { get; set; }
        public string MachineName { get; set; }
        public int ElapsedTime { get; set; }
        public string Platform { get; set; }
        public string SimulationType { get; set; }
        public string User { get; set; }
        public string[] EnvironmentTags { get; set; }
        public string KitDate { get; set; }
        public string FirstFail { get; set; }


    }
}
