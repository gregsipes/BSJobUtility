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

                List<string> files = GetFiles(sourceDirectory, new System.Text.RegularExpressions.Regex(@"[=]"));

                foreach (string file in files)
                {
                    FileInfo fileInfo = new FileInfo(file);

                    // if (fileInfo.LastWriteTime.Date == DateTime.Today.Date && !fileInfo.Name.Contains(".="))
                    if (!fileInfo.Name.Contains(".="))
                    {
                    //    WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.Name} last write time {fileInfo.LastAccessTime}");

                        Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PBS2Macro, "Proc_Select_Loads_If_Processed",
                                                                 new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                  new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();

                        if (result == null)
                        {
                            //make sure we the file is no longer being edited
                            if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                            {
                                //create new file name
                                string newFileName = fileInfo.Name.Replace("." + fileInfo.Extension, "") + "_" + DateTime.Now.ToString("yyyyMMddhhmmss tt") + ".txt";

                                WriteToJobLog(JobLogMessageType.INFO, $"Creating new loads record for {fileInfo.Name} last modified on  {new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)}");

                                //create load record
                                result = ExecuteSQL(DatabaseConnectionStringNames.PBS2Macro, "dbo.Proc_Insert_Loads",
                                                new SqlParameter("@pvchrOriginalDir", sourceDirectory),
                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                new SqlParameter("@pvchrUserName", Environment.UserName),
                                                new SqlParameter("@pvchrComputerName", Environment.MachineName),
                                                new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();

                                Int32 loadId = 0;
                                if (!Int32.TryParse(result["loads_id"].ToString(), out loadId))
                                {
                                    WriteToJobLog(JobLogMessageType.ERROR, result["loads_id"].ToString());
                                    throw new Exception(result["loads_id"].ToString());
                                }

                                if (loadId != 0)
                                {
                                    //copy file from source to destination
                                    File.Copy(file, destinationDirectory + newFileName);

                                    //update load record
                                    ExecuteNonQuery(DatabaseConnectionStringNames.PBS2Macro, "dbo.Proc_Update_Loads",
                                                    new SqlParameter("@pintLoadsID", loadId),
                                                    new SqlParameter("@pstrBackupFile", destinationDirectory + newFileName),
                                                    new SqlParameter("@plongFileSize", fileInfo.Length));

                                    WriteToJobLog(JobLogMessageType.INFO, "Copied " + file + " to " + destinationDirectory + newFileName);
                                }
                            } else
                                WriteToJobLog(JobLogMessageType.INFO, $"There's a chance the file is still getting updated, so we'll pick it up next run {fileInfo.Name}");


                           
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogException(ex);
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
