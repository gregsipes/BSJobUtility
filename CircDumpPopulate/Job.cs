using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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
            JobDescription = "Transfers data from a work (staging) database to a final table. This is step 2 in the PBS import process.";
            AppConfigSectionName = "CircDumpPopulate";
        }

        public override void ExecuteJob()
        {
            try
            {

                List<Dictionary<string, object>> dumpControls = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "Proc_Select_BN_Distinct_DumpControl_To_Populate",
                                                                               new SqlParameter("@pintGroupNumber", GroupNumber)).ToList();

                bool updateTranDateAfterSuccessfulPopulate = false;

                if (dumpControls != null && dumpControls.Count() > 0)
                {
                    foreach (Dictionary<string, object> dumpControl in dumpControls)
                    {
                        List<string> tablesToPopulate = DetermineTablesToPopulate(Convert.ToInt64(dumpControl["loads_dumpcontrol_id"]));

                        if (tablesToPopulate.Count() == 0)
                            WriteToJobLog(JobLogMessageType.INFO, $"No tables need to be populated for group number {GroupNumber}, loadds_dumpcontrol_id = {dumpControl["loads_dumpcontrol_id"].ToString()}");
                        else if (Convert.ToBoolean(dumpControl["flgUpdateTranDateAfterSuccessfulPopulate"].ToString()))
                            updateTranDateAfterSuccessfulPopulate = true;
                    }

                    //update tran date control file
                    if (updateTranDateAfterSuccessfulPopulate)
                        UpdateTranDateControlFile();
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
            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "Proc_Select_BN_Loads_Tables",
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
                        Dictionary<string, object> populateAttempts = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_Tables_Number_Of_Populate_Attempts",
                                                                                        new SqlParameter("@pbintLoadsTablesID", result["loads_tables_id"].ToString()),
                                                                                        new SqlParameter("@pintGroupNumber", GroupNumber)).FirstOrDefault();

                        if (Convert.ToInt32(populateAttempts["bn_loads_tables_number_of_populate_attempts"].ToString()) <= Convert.ToInt32(populateAttempts["bn_groups_number_of_populate_attempts"].ToString()))
                        {
                            if (Convert.ToBoolean(result["populate_error_flag"].ToString()))
                                WriteToJobLog(JobLogMessageType.INFO, $"{result["table_name"].ToString()} UNSUCCESSFUL in previous populate, but retrying populate");

                            //call to actually move the data
                            PopulateTable(Convert.ToInt64(result["loads_tables_id"].ToString()), result["table_name"].ToString(), Convert.ToBoolean(result["update_trannumber_control_file_after_populate_flag"].ToString()));

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

        private void PopulateTable(Int64 loadsTableId, string tableName, bool updateTranNumberControlFileAfterPopulate)
        {
            //todo:
        }
    }
}
