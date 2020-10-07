using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestListSynchronizer
{
    public class RecordUpdater : IRecordUpdater
    {
        private List<int> _incomingSuiteIDs = new List<int>();

        /// <summary>
        /// Called when a new record must be added top the database.  Occurs when a suite ID is in the
        /// input data, but is not present in the database.
        /// </summary>
        /// <param name="dbRecords">database records</param>
        /// <param name="xlRecords">excel records</param>
        public void NewRecord(IRecords dbRecords, IRecords xlRecords)
        {
            dbRecords.Add();
            dbRecords.SetFieldValue("Suite ID", SetValueOrDefault(xlRecords.GetFieldValue("Suite ID")));
            dbRecords.SetFieldValue("Suite Name", SetValueOrDefault(xlRecords.GetFieldValue("Name")));
            dbRecords.SetFieldValue("Defect", SetValueOrDefault(xlRecords.GetFieldValue("Defect")));
            dbRecords.SetFieldValue("Investigator", "unassigned");
            dbRecords.SetFieldValue("Machine", SetValueOrDefault(xlRecords.GetFieldValue("Machine")));
            dbRecords.SetFieldValue("Test Time", SetValueOrDefault(xlRecords.GetFieldValue("Test Time")));
            dbRecords.SetFieldValue("Result", SetValueOrDefault(xlRecords.GetFieldValue("Result")));
            dbRecords.SetFieldValue("Parent Result", "");
            dbRecords.SetFieldValue("Org", SetValueOrDefault(xlRecords.GetFieldValue("Org")));
            dbRecords.SetFieldValue("Platform", SetValueOrDefault(xlRecords.GetFieldValue("Platform")));
            dbRecords.SetFieldValue("Simulation", SetValueOrDefault(xlRecords.GetFieldValue("Simulation")));
            dbRecords.SetFieldValue("User", SetValueOrDefault(xlRecords.GetFieldValue("User")));
            dbRecords.SetFieldValue("Kit Type", SetValueOrDefault(xlRecords.GetFieldValue("Kit Type")));
            dbRecords.SetFieldValue("OS", SetValueOrDefault(xlRecords.GetFieldValue("OS")));
            dbRecords.SetFieldValue("Office", SetValueOrDefault(xlRecords.GetFieldValue("Office")));
            dbRecords.SetFieldValue("Kit", SetValueOrDefault(xlRecords.GetFieldValue("Kit")));
            dbRecords.SetFieldValue("Parent Kit", "");
            dbRecords.SetFieldValue("First Fail", SetValueOrDefault(xlRecords.GetFieldValue("First Fail")));
            dbRecords.SetFieldValue("Notes", "");
            dbRecords.SetFieldValue("Activity", "");
            dbRecords.SetFieldValue("Status", "");
            dbRecords.Update();
        }

        /// <summary>
        /// Called when an existing record in the database is updated with data from the input data.
        /// </summary>
        /// <param name="dbRecords">database records</param>
        /// <param name="xlRecords">excel records</param>
        public void UpdateRecord(IRecords dbRecords, IRecords xlRecords)
        {
            dbRecords.Edit();
            // We don't update Suite ID, Investigator, Notes, Status, or Activity
            // Investigator, Notes, Status, Activity do not come from results data.
            dbRecords.SetFieldValue("Suite Name", SetValueOrDefault(xlRecords.GetFieldValue("Name")));
            dbRecords.SetFieldValue("Defect", SetValueOrDefault(xlRecords.GetFieldValue("Defect")));
            dbRecords.SetFieldValue("Machine", SetValueOrDefault(xlRecords.GetFieldValue("Machine")));
            dbRecords.SetFieldValue("Test Time", SetValueOrDefault(xlRecords.GetFieldValue("Test Time")));
            dbRecords.SetFieldValue("Result", SetValueOrDefault(xlRecords.GetFieldValue("Result")));
            dbRecords.SetFieldValue("Org", SetValueOrDefault(xlRecords.GetFieldValue("Org")));
            dbRecords.SetFieldValue("Platform", SetValueOrDefault(xlRecords.GetFieldValue("Platform")));
            dbRecords.SetFieldValue("Simulation", SetValueOrDefault(xlRecords.GetFieldValue("Simulation")));
            dbRecords.SetFieldValue("User", SetValueOrDefault(xlRecords.GetFieldValue("User")));
            dbRecords.SetFieldValue("Kit Type", SetValueOrDefault(xlRecords.GetFieldValue("Kit Type")));
            dbRecords.SetFieldValue("OS", SetValueOrDefault(xlRecords.GetFieldValue("OS")));
            dbRecords.SetFieldValue("Office", SetValueOrDefault(xlRecords.GetFieldValue("Office")));
            dbRecords.SetFieldValue("Kit", SetValueOrDefault((xlRecords.GetFieldValue("Kit"))));
            dbRecords.SetFieldValue("First Fail", SetValueOrDefault(xlRecords.GetFieldValue("First Fail")));
            dbRecords.Update();
        }

        /// <summary>
        /// Called when updating record with parent data
        /// </summary>
        /// <param name="dbRecords">database records</param>
        /// <param name="xlRecords">excel records</param>
        public void UpdateParentRecord(IRecords dbRecords, IRecords xlRecords)
        {
            dbRecords.Edit();
            // When updating parent data we only need result and kit
            dbRecords.SetFieldValue("Parent Result", SetValueOrDefault(xlRecords.GetFieldValue("Result")));
            dbRecords.SetFieldValue("Parent Kit", SetValueOrDefault(xlRecords.GetFieldValue("Kit")));
            dbRecords.Update();
        }

        /// <summary>
        /// Called when checking existing suites to see if any were not run. If the suite exists in the database
        /// but was not in the input data then the Result field is set to Not Run.
        /// </summary>
        /// <param name="suiteID"></param>
        /// <param name="dbRecords"></param>
        public void NotRunRecord(int suiteID, IRecords dbRecords)
        {
            if (!_incomingSuiteIDs.Any(i => i == suiteID))
            {
                dbRecords.Edit();
                dbRecords.SetFieldValue("Result", "Not Run");
                dbRecords.Update();
            }
        }

        /// <summary>
        /// Track all incomming suiteIDs.
        /// </summary>
        /// <param name="suiteID"></param>
        public void AddIncommingRecord(int suiteID)
        {
            _incomingSuiteIDs.Add(suiteID);
        }

        /// <summary>
        /// Helper use when setting field values.  Hadles special case where strings must be " " for access 
        /// database.  Other wise the value of type default is legal.
        /// </summary>
        /// <param name="value">value to set the field</param>
        /// <returns>value or default</returns>
        public object SetValueOrDefault(object value)
        {
            switch (value)
            {
                // null or empty illegal as fields in db.
                case string s when string.IsNullOrEmpty(s):            
                    return " ";
                case null:
                    return " ";
                default:
                    return value;
            }
        }

        public void NewRecord(IRecords dbRecords, SuiteResult testResult)
        {
            dbRecords.Add();
            dbRecords.SetFieldValue("Suite ID", testResult.SuiteID);
            dbRecords.SetFieldValue("Suite Name", testResult.SuiteName);
            dbRecords.SetFieldValue("Defect", testResult.Defect);
            dbRecords.SetFieldValue("Investigator", "unassigned");
            dbRecords.SetFieldValue("Machine", testResult.MachineName);
            dbRecords.SetFieldValue("Test Time", testResult.ElapsedTime);
            dbRecords.SetFieldValue("Result", testResult.Result);
            dbRecords.SetFieldValue("Parent Result", "");
            dbRecords.SetFieldValue("Org", testResult.Organization);
            dbRecords.SetFieldValue("Platform", testResult.Platform);
            dbRecords.SetFieldValue("Simulation", testResult.SimulationType);
            dbRecords.SetFieldValue("User", testResult.User);
            dbRecords.SetFieldValue("Kit Type", testResult.EnvironmentTags[0]); // first element of array is kit type (asrt or bfr)
            dbRecords.SetFieldValue("OS", testResult.EnvironmentTags[1]); // second element of array is OS
            dbRecords.SetFieldValue("Office", testResult.EnvironmentTags[3]); //third element of array is office type
            dbRecords.SetFieldValue("Kit", testResult.KitDate);
            dbRecords.SetFieldValue("Parent Kit", "");
            dbRecords.SetFieldValue("First Fail", testResult.FirstFail);
            dbRecords.SetFieldValue("Notes", "");
            dbRecords.SetFieldValue("Activity", "");
            dbRecords.SetFieldValue("Status", "");
            dbRecords.Update();
        }

        public void UpdateRecord(IRecords dbRecords, SuiteResult testResult)
        {
            dbRecords.Edit();
            // We don't update Suite ID, Investigator, Notes, Status, or Activity
            // Investigator, Notes, Status, Activity do not come from results data.
            dbRecords.SetFieldValue("Suite Name", testResult.SuiteName);
            dbRecords.SetFieldValue("Defect", testResult.Defect);
            dbRecords.SetFieldValue("Machine", testResult.MachineName);
            dbRecords.SetFieldValue("Test Time", testResult.ElapsedTime);
            dbRecords.SetFieldValue("Result", testResult.Result);
            dbRecords.SetFieldValue("Org", testResult.Organization);
            dbRecords.SetFieldValue("Platform", testResult.Platform);
            dbRecords.SetFieldValue("Simulation", testResult.SimulationType);
            dbRecords.SetFieldValue("User", testResult.User);
            dbRecords.SetFieldValue("Kit Type", testResult.EnvironmentTags[0]); // first element of array is kit type (asrt or bfr)
            dbRecords.SetFieldValue("OS", testResult.EnvironmentTags[1]); // second element of array is OS
            dbRecords.SetFieldValue("Office", testResult.EnvironmentTags[3]); //third element of array is office type
            dbRecords.SetFieldValue("Kit", testResult.KitDate);
            dbRecords.SetFieldValue("First Fail", testResult.FirstFail);
            dbRecords.Update();
        }

        public void UpdateParentRecord(IRecords dbRecords, SuiteResult testResult)
        {
            dbRecords.Edit();
            // When updating parent data we only need result and kit
            dbRecords.SetFieldValue("Parent Result", testResult.Result);
            dbRecords.SetFieldValue("Parent Kit", testResult.KitDate);
            dbRecords.Update();
        }

    }
}
