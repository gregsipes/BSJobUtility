using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace AutoPrintPDF
{
    public class Job : JobBase
    {
        public string Version { get; set; }

        public override void SetupJob()
        {
            JobName = "AutoPrintPDF";
            JobDescription = "TODO";
            AppConfigSectionName = "AutoPrintPDF";

        }

        public override void ExecuteJob()
        {
            try
            {
                switch (Version)
                {
                    case "OfficePay":
                    case "AutoRenew":
                        //todo:
                        break;
                    case "PBSInvoices":
                        //todo:
                        break;
                    case "PBSInvoicesByCarrierID":
                        //todo:
                        break;
                    default:
                        throw new Exception("Unknown version");
                }
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void AutoRenewOrOfficePay()
        {
            string description = "renewal";

            if (Version == "AutoRenew")
                description = "autorenew";

            WriteToJobLog(JobLogMessageType.INFO, $"Determining {description} notices to send to .pdf");

            List<Dictionary<string, object>> loads = new List<Dictionary<string, object>>();

            if (Version == "AutoRenew")
                loads = ExecuteSQL(DatabaseConnectionStringNames.AutoRenew, "Proc_Select_Loads_For_AutoPrint_To_PDF").ToList();
            else
                loads = ExecuteSQL(DatabaseConnectionStringNames.OfficePay, "Proc_Select_Loads_For_AutoPrint_To_PDF").ToList();

            if (loads == null || loads.Count() == 0)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"No {description} notices to create .pdf's for exist in database");
                return;
            }

            //todo: install tru type font?


            WriteToJobLog(JobLogMessageType.INFO, "Creating .pdf's in work directory as {subscription_number}{MMddyyyy}INVOICE.pdf");

            foreach (Dictionary<string, object> load in loads)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Retrieving {description} notices for loads_id {load["loads_id"].ToString()}");

                List<Dictionary<string, object>> results = new List<Dictionary<string, object>>();

                if (Version == "AutoRenew")
                    results = ExecuteSQL(DatabaseConnectionStringNames.AutoRenew, "Proc_Select_For_AutoRenew",
                                                        new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                                        new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()),
                                                        new SqlParameter("@pflgOnlyWithEmailAddress", false),
                                                        new SqlParameter("@pflgReport", true),
                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword"))).ToList();
                else
                    results = ExecuteSQL(DatabaseConnectionStringNames.OfficePay, "Proc_Select_For_Office_Pay_Bills",
                                                        new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                                        new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()),
                                                        new SqlParameter("@pvchrRenewalType", null),
                                                        new SqlParameter("@pvchrRenewalNumber", 0),
                                                        new SqlParameter("@pflgZero", 0),
                                                        new SqlParameter("@pflgDuplicate", 0),
                                                        new SqlParameter("@pflgAutoPrintToPDF", 1),
                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword"))).ToList();


                if (results == null || results.Count() == 0)
                {
                    WriteToJobLog(JobLogMessageType.INFO, $"No {description} notices exist for this loads_id");
                    return;
                }

                string outputDirectory = GetConfigurationKeyValue("WorkDirectory") + Version;

            }


        }

    }
}
