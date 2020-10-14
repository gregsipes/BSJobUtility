using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace PBSDumpPopulate
{
    public class Job : JobBase
    {
        public string Version { get; set; }
        public string GroupName { get; set; }

        public override void SetupJob()
        {
            JobName = "PBS Dump Populate";
            JobDescription = "Transfers data from a work (staging) database to a final table (PBSDump). This is step 2 in the PBS import process. This is only for PBSDumpA since the B and C versions populate at the end of the workload/first step";
            AppConfigSectionName = "PBSDumpPopulate";
        }

        public override void ExecuteJob()
        {
            try
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Group Name: {GroupName}   Group Number: {Version}");

                List<Dictionary<string, object>> dumpControls = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Select_BN_Distinct_DumpControl_To_Populate",
                                                                               new SqlParameter("@pintGroupNumber", Version)).ToList();

                //  bool updateTranDateAfterSuccessfulPopulate = false;

                if (dumpControls != null && dumpControls.Count() > 0)
                {
                    foreach (Dictionary<string, object> dumpControl in dumpControls)
                    {
                        List<string> tablesToPopulate = DetermineTablesToPopulate(Convert.ToInt64(dumpControl["loads_dumpcontrol_id"]));

                        if (tablesToPopulate.Count() == 0)
                            WriteToJobLog(JobLogMessageType.INFO, $"No tables need to be populated for group number {Version}, loads_dumpcontrol_id = {dumpControl["loads_dumpcontrol_id"].ToString()}");

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSDumpAWorkLoad, "Proc_Update_BN_Loads_DumpControl_Load_Successful_Flag",
                                                new SqlParameter("@pintLoadsDumpControlID", dumpControl["loads_dumpcontrol_id"]));
                        //else if (Convert.ToBoolean(dumpControl["flgUpdateTranDateAfterSuccessfulPopulate"].ToString()))
                        //    updateTranDateAfterSuccessfulPopulate = true; //this is never true
                    }

                    //update tran date control file
                    //  if (updateTranDateAfterSuccessfulPopulate)
                    //UpdateTranDateControlFile();
                }




            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private List<string> DetermineTablesToPopulate(Int64 dumpControlId)
        {
            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Select_BN_Loads_Tables",
                                                                                new SqlParameter("@pbintLoadsDumpControlID", dumpControlId)).ToList();

            List<string> tablesToUpdate = new List<string>();

            if (results != null && results.Count() > 0)
            {
                //todo: send executing email?

                WriteToJobLog(JobLogMessageType.INFO, $"Populating tables for group number {Version}, loads dump control id = {dumpControlId}");

                foreach (Dictionary<string, object> result in results)
                {
                    if (Convert.ToBoolean(result["populate_successful_flag"].ToString()))
                        WriteToJobLog(JobLogMessageType.INFO, $"{result["table_name"].ToString()} successful in previous populate");
                    else
                    {
                        Dictionary<string, object> populateAttempts = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Update_BN_Loads_Tables_Number_Of_Populate_Attempts",
                                                                                        new SqlParameter("@pbintLoadsTablesID", result["loads_tables_id"].ToString()),
                                                                                        new SqlParameter("@pintGroupNumber", Version)).FirstOrDefault();

                        if (Convert.ToInt32(populateAttempts["bn_loads_tables_number_of_populate_attempts"].ToString()) <= Convert.ToInt32(populateAttempts["bn_groups_number_of_populate_attempts"].ToString()))
                        {
                            if (Convert.ToBoolean(result["populate_error_flag"].ToString()))
                                WriteToJobLog(JobLogMessageType.INFO, $"{result["table_name"].ToString()} UNSUCCESSFUL in previous populate, but retrying populate");

                            //call to actually move the data
                            PopulateTable(Convert.ToInt64(result["loads_tables_id"].ToString()), result["table_name"].ToString());

                            tablesToUpdate.Add(result["table_name"].ToString());
                        }
                        else
                        {
                            if (Convert.ToInt32(populateAttempts["bn_loads_tables_number_of_populate_attempts"].ToString()) > Convert.ToInt32(populateAttempts["bn_groups_number_of_populate_attempts"].ToString()))
                            {
                                WriteToJobLog(JobLogMessageType.WARNING, $"{result["table_name"].ToString()} UNSUCCESSFUL in previous populate");

                                //todo: send email?
                            }
                        }
                    }
                }
            }


            return tablesToUpdate;

        }

        private void PopulateTable(Int64 loadsTableId, string tableName)
        {
            bool hasError = false;
            bool perge = true;
            bool verify = false;

            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Select_BN_Tables",
                                                        new SqlParameter("@pvchrTableName", tableName)).FirstOrDefault();

            if (result != null && result.Count() > 0)
            {
                perge = Convert.ToBoolean(result["purge_flag"].ToString());
                verify = Convert.ToBoolean(result["verify_flag"].ToString());
            }

            if (perge)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Purging {tableName}");
                result = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Purge_Table",
                                                new SqlParameter("@pbintLoadsTablesID", loadsTableId)).FirstOrDefault();

                if (result["error_message"].ToString() != "")
                {
                    WriteToJobLog(JobLogMessageType.ERROR, result["error_message"].ToString());
                    hasError = true;
                }
            }
            else
                WriteToJobLog(JobLogMessageType.INFO, $"{tableName} bypassing purge");

            WriteToJobLog(JobLogMessageType.INFO, $"{tableName} populating");
            result = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, $"Proc_Populate_{tableName}",
                                                                    new SqlParameter("@pbintLoadsTablesID", loadsTableId)).FirstOrDefault();

            if (result["error_message"].ToString() != "")
            {
                hasError = true;
                WriteToJobLog(JobLogMessageType.ERROR, result["error_message"].ToString());
            }


            if (verify)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Verifying {tableName}");

                result = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Verify_Table",
                                new SqlParameter("@pbintLoadsTablesID", loadsTableId)).FirstOrDefault();

                if (result["error_message"].ToString() != "")
                {
                    WriteToJobLog(JobLogMessageType.ERROR, result["error_message"].ToString());
                    hasError = true;
                }
            }
            else
            {
                WriteToJobLog(JobLogMessageType.INFO, $"{tableName} bypassing verify");
                WriteToJobLog(JobLogMessageType.INFO, $"{tableName} delete work records");

                result = ExecuteSQL(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Delete_Table",
                                                           new SqlParameter("@pbintLoadsTablesID", loadsTableId)).FirstOrDefault();

                if (result["error_message"].ToString() != "")
                {
                    hasError = true;
                    WriteToJobLog(JobLogMessageType.ERROR, result["error_message"].ToString());
                }
            }



            if (!hasError)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Update_BN_Loads_Tables_Populate_Successful_Flag",
                                                        new SqlParameter("@pbintLoadsTablesID", loadsTableId));

                ////delete unsuccessful touch file if one exists
                //if (File.Exists($"{GetConfigurationKeyValue("TableTouchDirectory")}{GroupName}\\{tableName}.unsuccessful"))
                //    File.Delete($"{GetConfigurationKeyValue("TableTouchDirectory")}{GroupName}\\{tableName}.unsuccessful");

                ////create a successul file (this is the file that gets cleaned up in the next step of the process (CircDumpPost)
                //File.Create($"{GetConfigurationKeyValue("TableTouchDirectory")}{GroupName}\\{tableName}.successful");

            }
            else
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.PBSDumpAWorkPopulate, "Proc_Update_BN_Loads_Tables_Populate_Error_Flag",
                                        new SqlParameter("@pbintLoadsTablesID", loadsTableId));

                SendMail($"{JobName} - Error", result["error_message"].ToString(), false);
            }

        }
    }
}
