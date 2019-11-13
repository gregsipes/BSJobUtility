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
                    SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.PBS2Macro, CommandType.StoredProcedure, "dbo.Proc_Insert_Loads",
                                     new Dictionary<string, object>()
                                                         {
                                                            { "@pvchrOriginalDir", sourceDirectory },
                                                            { "@pvchrOriginalFile", fileInfo.Name },
                                                            { "@pdatLastModified", fileInfo.LastWriteTime },
                                                            { "@pvchrUserName", Environment.UserName },
                                                            { "@pvchrComputerName", Environment.MachineName } ,
                                                            { "@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString() }
                                                         });

                    Int32? loadId = reader.GetInt32(reader.GetOrdinal("loads_id"));

                    if (loadId == null)
                    {
                        //copy file from source to destination
                        File.Copy(file, destinationDirectory + newFileName);


                        //update load record
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBS2Macro, CommandType.StoredProcedure, "dbo.Proc_Update_Loads",
                                        new Dictionary<string, object>()
                                        {
                                            { "@pintLoadsID", loadId },
                                            { "@pstrBackupFile", destinationDirectory + newFileName }
                                        });

                        WriteToJobLog(JobLogMessageType.INFO, "Copied " + file + " to " + destinationDirectory + newFileName);
                    }
                }
            }
            catch (Exception)
            {
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
