using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Newtonsoft.Json;

namespace TestListSynchronizer
{
    internal static class JarvisWrapper
    {
        static string fields = "--oSuiteId --oResult --oOrganization --oSuiteName --oDefect --oMachineName --oElapsedTime --oPlatform --oSimulationType --oUser --oEnvironmentTags --oKitDate --oFirstFail";

        public static List<SuiteResult> FetchResults(string project)
        {
            return JsonConvert.DeserializeObject<List<SuiteResult>>(JarvisQueryRawResults(project, null));
        }

        public static List<SuiteResult> FetchResults(string project, string baseline)
        {
            return JsonConvert.DeserializeObject<List<SuiteResult>>(JarvisQueryRawResults(project, baseline));
        }

        private static string JarvisQueryRawResults(string project, string baseline)
        {
            string jarvisApp = @"C:\Users\hlotyaks\bin\Jarvis.exe";

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
