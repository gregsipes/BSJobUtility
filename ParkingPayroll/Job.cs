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
                ExecuteNonQuery(DatabaseConnectionStringNames.Parking, sproc, 
                                new SqlParameter("@pvchrServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                new SqlParameter("@pvchrDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")));

                WriteToJobLog(JobLogMessageType.INFO, "Executed " + sproc + " on " + DatabaseConnectionStringNames.Parking.ToString());
            }
            catch (Exception ex)
            {
                SendMail($"Error in Job: {JobName}", ex.ToString(), false);
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
