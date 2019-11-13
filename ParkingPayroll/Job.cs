using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParkingPayroll
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {
                string sproc = "dbo.Proc_Insert_Update_People_From_Payroll";
                ExecuteNonQuery(DatabaseConnectionStringNames.Parking, CommandType.StoredProcedure, sproc, 
                                new Dictionary<string, object>() {
                                        { "@pvchrServerInstance", GetConfigurationKeyValue("RemoteServerInstance") },
                                        { "@pvchrDatabase", GetConfigurationKeyValue("RemoteDatabaseName") },
                                        { "@pvchrUserName", GetConfigurationKeyValue("RemoteUserName") },
                                        { "@pvchrPassword", GetConfigurationKeyValue("RemotePassword") }
                                    });

                WriteToJobLog(JobLogMessageType.INFO, "Executed " + sproc + " on " + DatabaseConnectionStringNames.Parking.ToString());
            }
            catch (Exception)
            {

                throw;
            }

        }

        public override void SetupJob()
        {
            JobName = "Parking Payroll";
            JobDescription = "Updates employees and departments from SBS";
            AppConfigSectionName = "ParkingPayroll";
        }
    }
}
