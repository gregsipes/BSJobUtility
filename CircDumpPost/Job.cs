using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace CircDumpPost
{
    public class Job : JobBase
    {
        public int GroupNumber { get; set; }

        public override void SetupJob()
        {
            JobName = "Circ Dump Post";
            JobDescription = "Runs final cleanup and update tasks for circ dump.";
            AppConfigSectionName = "CircDumpPost";

        }

        public override void ExecuteJob()
        {
            try
            {
                if (GroupNumber > 0)
                    GroupPost();
                else
                    TablePost();
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void GroupPost()
        {
            //only run this code if a successful file exists. This files gets deleted after the table post  (last) step of the job is complete
            List<string> files = Directory.GetFiles($"{GetConfigurationKeyValue("TableTouchDirectory")}", "*.successful").ToList();

            if (files.Count() > 0)
            {
                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpPost, "Proc_Select_BN_Groups_Post_Load",
                                                                            new SqlParameter("@pintGroupNumber", GroupNumber));

                if (GroupNumber == -1)
                    WriteToJobLog(JobLogMessageType.INFO, $"Preparing to execute {results.Count()} post-load routines for all tables");
                else
                    WriteToJobLog(JobLogMessageType.INFO, $"Preparing to execute {results.Count()} post-load routines for group number {GroupNumber}");

                Exception sprocResult = null;
                foreach (Dictionary<string, object> result in results)
                {
                    sprocResult = ExecuteStoredProcedure(true, Convert.ToInt64(result["bn_groups_post_load_id"].ToString()), result["stored_procedure"].ToString(), "", Convert.ToInt32(result["database_number"].ToString()));

                    //if something went wrong, determine if the job should continue processing or exit
                    if (sprocResult != null)
                    {
                        if (!Convert.ToBoolean(result["continue_on_failure_flag"].ToString()))
                            throw new Exception(sprocResult.ToString());
                    }
                }
            }

        }

        private Exception ExecuteStoredProcedure(bool isGroup, Int64 postLoadId, string sproc, string tableName, Int32 databaseNumber)
        {
            try
            {
                string parameters = RetrieveParameters(isGroup, postLoadId, tableName, true);
                WriteToJobLog(JobLogMessageType.INFO, $"{sproc} executing");

                DatabaseConnectionStringNames database;

                switch (databaseNumber)
                {
                    case 2:
                        database = DatabaseConnectionStringNames.PBSDump;
                        break;
                    case 3:
                        database = DatabaseConnectionStringNames.BNTransactions;
                        break;
                    default:
                        database = DatabaseConnectionStringNames.CircDumpWorkLoad;
                        break;
                }

                ExecuteNonQuery(database, CommandType.Text, "EXEC " + sproc + " " + parameters);

                return null;
            }
            catch (Exception ex)
            {
                WriteToJobLog(JobLogMessageType.ERROR, ex.ToString());
                return ex;
            }
        }

        private string RetrieveParameters(bool isGroup, Int64 postLoadId, string tableName, bool quote)
        {

            List<Dictionary<string, object>> results = new List<Dictionary<string, object>>();

            if (isGroup)
                results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpPost, "Proc_Select_BN_Groups_Post_Load_Parameters", new SqlParameter("@pintBNGroupsPostLoadID", postLoadId)).ToList();
            else
                results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpPost, "Proc_Select_BN_Tables_Post_Load_Parameters", new SqlParameter("@pintBNTablesPostLoadID", postLoadId)).ToList();

            string parameterString = "";

            foreach (Dictionary<string, object> result in results)
            {
                switch (result["parameter_name"].ToString())
                {
                    case "pbsdump_server_instance":
                        parameterString += "'" + GetConfigurationKeyValue("RemoteServerInstance") + "'" + ",";
                        break;
                    case "pbsdump_database":
                        parameterString += "'" + GetConfigurationKeyValue("RemoteDatabaseName") + "'" + ",";
                        break;
                    case "user_name":
                        parameterString += "'" + GetConfigurationKeyValue("RemoteUserName") + "'" + ",";
                        break;
                    case "password":
                        parameterString += "'" + GetConfigurationKeyValue("RemotePassword") + "'";
                        break;
                }

            }

            return parameterString;
        }

        private void TablePost()
        {

            List<string> files = Directory.GetFiles($"{GetConfigurationKeyValue("TableTouchDirectory")}", "*.successful").ToList();

            foreach (string file in files)
            {
                FileInfo fileInfo = new FileInfo(file);

                string tableName = fileInfo.Name.Substring(0, fileInfo.Name.LastIndexOf("."));

                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpPost, "Proc_Select_BN_Tables_Post_Load",
                                                                               new SqlParameter("@pvchrTableName", tableName));

                WriteToJobLog(JobLogMessageType.INFO, $"Preparing to execute {results.Count()} post-load routines for group number {GroupNumber}");

                Exception sprocResult = null;
                foreach (Dictionary<string, object> result in results)
                {
                    if (result["stored_procedure"].ToString() != null && result["stored_procedure"].ToString() != "")
                        sprocResult = ExecuteStoredProcedure(false, Convert.ToInt64(result["bn_tables_post_load_id"].ToString()), result["stored_procedure"].ToString(), tableName, Convert.ToInt32(result["database_number"].ToString())); //execute sproc
                   //  else
                        //  ExecuteExecutable();   //execute INI file? This doesn't appear to be in use any longer


                    //if something went wrong, determine if the job should continue processing or exit
                    if (sprocResult != null)
                    {
                        if (!Convert.ToBoolean(result["continue_on_failure_flag"].ToString()))
                            throw new Exception(sprocResult.ToString());
                    }
                }

                File.Delete(file);
            }
        }
    }
}
