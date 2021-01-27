using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace ExecuteSQL
{
    public class Job : JobBase
    {
        public string Version { get; set; }

        public override void SetupJob()
        {
            JobName = "Execute SQL - " + Version;
            JobDescription = "Executes SQL sprocs against different production databases";
            AppConfigSectionName = "ExecuteSQL";
        }

        public override void ExecuteJob()
        {
            WriteToJobLog(JobLogMessageType.INFO, $"Executing stored procedures...");

            try
            {
                switch (Version)
                {
                    case "BNTransactions":
                        ExecuteNonQuery(DatabaseConnectionStringNames.BNTransactions, "Proc_Delete_PBSComplaints_New_IDs_Older_Than");
                        break;
                    case "BSConsole_EndOfMonth":
                        ExecuteNonQuery(DatabaseConnectionStringNames.BSConsole, "Proc_Insert_Launches_Historical");
                        break;
                    case "Brainworks_ArchiveHistory_BWDB_BW":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Brainworks, "ArchiveHistory");
                        break;
                    case "Brainworks_CMPromiseToPay_BWDB_BW":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Brainworks, "CMPromiseToPay");
                        break;
                    case "Brainworks_Previous_Last_Posted":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Brainworks, "Proc_BuffNews_Update_tblCustomFieldsValues_Prev_Last_Posted");
                        break;
                    case "Brainworks_procAutoTransferAdjustment_BWDB_BW":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Brainworks, "procAutoTransferAdjustment");
                        break;
                    case "Brainworks_update_tblMovements_BWDB_BW":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Brainworks, "update_tblMovements");
                        break;
                    case "Brainworks_Work":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Brainworks_Work, "Proc_Delete_Ins_Work");
                        break;
                    case "BSConsole":
                        ExecuteNonQuery(DatabaseConnectionStringNames.BSConsole, "Proc_Delete_Launches");
                        break;
                    case "BuffNewsForBW":
                        ExecuteNonQuery(DatabaseConnectionStringNames.BuffNewsForBW, "Proc_Insert_Delete_Transactions_Credits_Debits",
                                                            new SqlParameter("@pvchrServerInstance", GetConfigurationKeyValue("BrainworksServer")),
                                                            new SqlParameter("@pvchrDatabase", GetConfigurationKeyValue("BrainworksDatabase")),
                                                            new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("BrainworksUserName")),
                                                            new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("BrainworksPassword")),
                                                            new SqlParameter("@pvchrLastModifiedBy", System.Security.Principal.WindowsIdentity.GetCurrent().Name));

                        ExecuteNonQuery(DatabaseConnectionStringNames.BuffNewsForBW, "Proc_Insert_Delete_Transactions_Flash",
                                                            new SqlParameter("@pvchrServerInstance", GetConfigurationKeyValue("BrainworksServer")),
                                                            new SqlParameter("@pvchrDatabase", GetConfigurationKeyValue("BrainworksDatabase")),
                                                            new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("BrainworksUserName")),
                                                            new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("BrainworksPassword")),
                                                            new SqlParameter("@pvchrLastModifiedBy", System.Security.Principal.WindowsIdentity.GetCurrent().Name));

                        ExecuteNonQuery(DatabaseConnectionStringNames.BuffNewsForBW, "Proc_Insert_Delete_Transactions_Divisions",
                                                            new SqlParameter("@pvchrServerInstance", GetConfigurationKeyValue("BrainworksServer")),
                                                            new SqlParameter("@pvchrDatabase", GetConfigurationKeyValue("BrainworksDatabase")),
                                                            new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("BrainworksUserName")),
                                                            new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("BrainworksPassword")),
                                                            new SqlParameter("@pvchrLastModifiedBy", System.Security.Principal.WindowsIdentity.GetCurrent().Name));
                        break;
                    case "CircDump_Work":
                        ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Delete_BN_Loads_DumpControl");
                        break;
                    case "CommissionsRelated_Responsible_Salespersons_BWDB_BW":
                        ExecuteNonQuery(DatabaseConnectionStringNames.CommissionsRelated, "Proc_Insert_Update_Responsible_Salespersons");
                        break;
                    case "MailTops":
                        ExecuteNonQuery(DatabaseConnectionStringNames.MailTops, "Proc_Purge_Trucks");
                        break;
                    case "Newshole":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Newshole, "Proc_Insert_Brainworks",
                                                            new SqlParameter("@pvchrBrainworksServiceInstance", GetConfigurationKeyValue("BrainworksServer")),
                                                            new SqlParameter("@pvchrBrainworksDatabase", GetConfigurationKeyValue("BrainworksDatabase")),
                                                            new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("BrainworksUserName")),
                                                            new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("BrainworksPassword")));
                        break;
                    case "OfficePay_Archived":
                        ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay_Archived, "Proc_Insert_Archived");
                        break;
                    case "Palletizers":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Palletizers, "Proc_Purge_Imports");
                        break;
                    case "PBSInvoiceExport":
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExport, "Proc_Delete_Work_IDs_Old");
                        break;
                    case "PBSInvoices":
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "Proc_Delete_Sessions_Older_Than_1_Day");
                        break;
                    case "Postings":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Postings, "Proc_Reset_Inquiries");
                        break;
                    case "Preprints":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Preprints, "Proc_Insert_Update_Brainworks_Related", 
                                                            new SqlParameter("@pintBrainworksAccount", "0"),
                                                            new SqlParameter("@pvchrBrainworksServiceInstance", GetConfigurationKeyValue("BrainworksServer")),
                                                            new SqlParameter("@pvchrBrainworksDatabase", GetConfigurationKeyValue("BrainworksDatabase")),
                                                            new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("BrainworksUserName")),
                                                            new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("BrainworksPassword")));
                        break;
                    case "Preprints_Archive":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Preprints, "Proc_Archive_Scheduled");
                        break;
                    case "Supplies_Work":
                        ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Delete_BN_Loads_DumpControl");
                        break;
                    case "TouchControl":
                        ExecuteNonQuery(DatabaseConnectionStringNames.TouchControl, "Proc_Delete_Outcomes_Files");
                        ExecuteNonQuery(DatabaseConnectionStringNames.TouchControl, "Proc_Delete_Schedulables_Creations");
                        break;
                    case "TouchControl_Delete_TouchControlService_Activity":
                        ExecuteNonQuery(DatabaseConnectionStringNames.TouchControl, "Proc_Delete_TouchControlService_Activity");
                        break;
                    case "Trade":
                        ExecuteNonQuery(DatabaseConnectionStringNames.Trade, "Proc_Update_Requested_Closed");
                        break;
                    default:
                        throw new Exception("Version not supported");

                }

                WriteToJobLog(JobLogMessageType.INFO, $"Successfully executed procedures");

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

    }
}
