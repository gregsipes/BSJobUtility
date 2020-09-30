using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace CircDumpPopulate
{
    public class Job : JobBase
    {

        public int GroupNumber { get; set; }

        public override void SetupJob()
        {
            JobName = "Circ Dump Populate";
            JobDescription = "Transfers data from a work (staging) database to a final table (PBSDump). This is step 2 in the PBS import process.";
            AppConfigSectionName = "CircDumpPopulate";
        }

        public override void ExecuteJob()
        {
            try
            {

                List<Dictionary<string, object>> dumpControls = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkPopulate, "Proc_Select_BN_Distinct_DumpControl_To_Populate",
                                                                               new SqlParameter("@pintGroupNumber", GroupNumber)).ToList();

              //  bool updateTranDateAfterSuccessfulPopulate = false;

                if (dumpControls != null && dumpControls.Count() > 0)
                {
                    foreach (Dictionary<string, object> dumpControl in dumpControls)
                    {
                        List<string> tablesToPopulate = DetermineTablesToPopulate(Convert.ToInt64(dumpControl["loads_dumpcontrol_id"]));

                        if (tablesToPopulate.Count() == 0)
                            WriteToJobLog(JobLogMessageType.INFO, $"No tables need to be populated for group number {GroupNumber}, loadds_dumpcontrol_id = {dumpControl["loads_dumpcontrol_id"].ToString()}");
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
            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkPopulate, "Proc_Select_BN_Loads_Tables",
                                                                                new SqlParameter("@pbintLoadsDumpControlID", dumpControlId)).ToList();

            List<string> tablesToUpdate = new List<string>();

            if (results != null && results.Count() > 0)
            {
                //todo: send executing email?

                foreach (Dictionary<string, object> result in results)
                {
                    if (Convert.ToBoolean(result["populate_successful_flag"].ToString()))
                        WriteToJobLog(JobLogMessageType.INFO, $"{result["table_name"].ToString()} successful in previous populate");
                    else
                    {
                        Dictionary<string, object> populateAttempts = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkPopulate, "Proc_Update_BN_Loads_Tables_Number_Of_Populate_Attempts",
                                                                                        new SqlParameter("@pbintLoadsTablesID", result["loads_tables_id"].ToString()),
                                                                                        new SqlParameter("@pintGroupNumber", GroupNumber)).FirstOrDefault();

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

            WriteToJobLog(JobLogMessageType.INFO, $"{tableName} bypassing purge");

            WriteToJobLog(JobLogMessageType.INFO, $"{tableName} populating");
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkPopulate, $"Proc_Populate_{tableName}",
                                                                    new SqlParameter("@pbintLoadsTablesID", loadsTableId)).FirstOrDefault();

            if (result["error_message"].ToString() != "")
            {
                hasError = true;
                WriteToJobLog(JobLogMessageType.ERROR, result["error_message"].ToString());
            }


            WriteToJobLog(JobLogMessageType.INFO, $"{tableName} bypassing verify");
            WriteToJobLog(JobLogMessageType.INFO, $"{tableName} delete work records");

             result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkPopulate, "Proc_Delete_Table",
                                                        new SqlParameter("@pbintLoadsTablesID", loadsTableId)).FirstOrDefault();

            if (result["error_message"].ToString() != "")
            {
                hasError = true;
                WriteToJobLog(JobLogMessageType.ERROR, result["error_message"].ToString());
            }


            if (!hasError)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkPopulate, "Proc_Update_BN_Loads_Tables_Populate_Successful_Flag",
                                                        new SqlParameter("@pbintLoadsTablesID", loadsTableId));

                //delete unsuccessful touch file if one exists
                //if (File.Exists($"{GetConfigurationKeyValue("TableTouchDirectory")}{GroupNumber}\\{tableName}.unsuccessful"))
                //    File.Delete($"{GetConfigurationKeyValue("TableTouchDirectory")}{GroupNumber}\\{tableName}.unsuccessful");

                ////create a successul file (this is the file that gets cleaned up in the next step of the process (CircDumpPost)
                //File.Create($"{GetConfigurationKeyValue("TableTouchDirectory")}{GroupNumber}\\{tableName}.successful");

            }
            else
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkPopulate, "Proc_Update_BN_Loads_Tables_Populate_Error_Flag",
                                        new SqlParameter("@pbintLoadsTablesID", loadsTableId));

                SendMail($"{JobName} - Error", result["error_message"].ToString(), false);
            }

        }


    }
}
