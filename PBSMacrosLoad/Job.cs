using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Reflection;

namespace PBSMacrosLoad
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {
                string sourceDirectory = GetConfigurationKeyValue("SourceDirectory");
                string destinationDirectory = GetConfigurationKeyValue("DestinationDirectory");

                foreach (string file in GetFiles(sourceDirectory, new System.Text.RegularExpressions.Regex(@"[=]")))
                {
                    FileInfo fileInfo = new FileInfo(file);

                    //create new file name
                    string newFileName = fileInfo.Name.Replace("." + fileInfo.Extension, "") + "_" + DateTime.Now.ToString("yyyyMMddhhmmss tt") + ".txt";

                    //create load record
                    Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PBS2Macro, "dbo.Proc_Insert_Loads",
                                     new SqlParameter("@pvchrOriginalDir", sourceDirectory),
                                     new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                     new SqlParameter("@pdatLastModified", fileInfo.LastWriteTime),
                                     new SqlParameter("@pvchrUserName", Environment.UserName),
                                     new SqlParameter("@pvchrComputerName", Environment.MachineName),
                                     new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
                
                    Int32? loadId = (Int32)result["loads_id"];

                    if (loadId == null)
                    {
                        //copy file from source to destination
                        File.Copy(file, destinationDirectory + newFileName);


                        //update load record
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBS2Macro, "dbo.Proc_Update_Loads",
                                        new SqlParameter("@pintLoadsID", loadId),
                                        new SqlParameter("@pstrBackupFile", destinationDirectory + newFileName));
                    
                        WriteToJobLog(JobLogMessageType.INFO, "Copied " + file + " to " + destinationDirectory + newFileName);
                    }
                }
            }
            catch (Exception ex)
            {
                SendMail($"Error in Job: {JobName}", ex.ToString(), false);
                throw;
            }
        }

        public override void SetupJob()
        {
            JobName = "PBS Macro Load";
            JobDescription = @"Moves files from \\circ\spoolcmro to \\Synergy\SERops\To Be Loaded\PBS ASCII";
            AppConfigSectionName = "PBSMacroLoad";
        }

    }
}
