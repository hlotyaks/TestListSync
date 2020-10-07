using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Newtonsoft.Json;

namespace TestListSynchronizer
{
    public class JarvisWrapper :IJarvisWrapper
    {
        static string fields = "--oSuiteId --oResult --oOrganization --oSuiteName --oDefect --oMachineName --oElapsedTime --oPlatform --oSimulationType --oUser --oEnvironmentTags --oKitDate --oFirstFail";
        string _jarvisApp;

        public JarvisWrapper(string jarvisApp)
        {
            _jarvisApp = jarvisApp;
        }

        public List<SuiteResult> FetchResults(string project)
        {
            return JsonConvert.DeserializeObject<List<SuiteResult>>(JarvisQueryRawResults(project, null, _jarvisApp));
        }

        public List<SuiteResult> FetchResults(string project, string baseline)
        {
            return JsonConvert.DeserializeObject<List<SuiteResult>>(JarvisQueryRawResults(project, baseline, _jarvisApp));
        }

        private static string JarvisQueryRawResults(string project, string baseline, string jarvisApp)
        {
            string args = (baseline == null) ?
                $"queryRawResults --project {project} {fields}" :
                $"queryRawResults --project {project} --baseline {baseline} {fields}";

            var jarvisProc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = jarvisApp,
                    Arguments = args,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                }
            };

            jarvisProc.Start();

            return jarvisProc.StandardOutput.ReadToEnd();
        }
    }
}
