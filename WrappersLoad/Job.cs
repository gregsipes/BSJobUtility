using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace WrappersLoad
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "labelsA*").ToList();



                if (files != null && files.Count() > 0)
                {
                    List<Dictionary<string, object>> previouslyLoadedFiles = ExecuteSQL(DatabaseConnectionStringNames.Wrappers, "dbo.Proc_Select_Loads").ToList();

                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> loadedFile = previouslyLoadedFiles.Where(p => p["original_file"].ToString() == fileInfo.Name).OrderByDescending(p => p["loads_id"]).FirstOrDefault();
                        

                        if ((DateTime.Now - fileInfo.LastWriteTime).TotalDays <= 1 && (loadedFile == null || DateTime.Compare(fileInfo.LastWriteTime.Date, DateTime.Parse(loadedFile["original_file_last_modified"].ToString()).Date) > 0))
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                            CopyAndProcessFile(fileInfo);
                        }
                        //else
                        //{
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.Wrappers, "Proc_Insert_Loads_Not_Loaded",
                        //                    new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString()),
                        //                    new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                        //                    new SqlParameter("@pdatLastModified", fileInfo.LastWriteTime),
                        //                    new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                        //                    new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                        //                    new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString()));
                        //}
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
            JobName = "Wrappers Load";
            JobDescription = "Builds wrapper cover/label to be printed";
            AppConfigSectionName = "WrappersLoad";
        }

        private void CopyAndProcessFile(FileInfo fileInfo)
        {
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Wrappers, "Proc_Insert_Loads",
                                                                                       new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString()),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", fileInfo.LastWriteTime),
                                                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.Wrappers, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            Int32 groupCounter = 0;
            Int32 groupLineCounter = 0;
            foreach (string line in fileContents)
            {
                if (line != null && line.Trim().Length > 0)
                {
                    if (line.Trim().StartsWith("*****"))
                    {
                        groupCounter++;
                        groupLineCounter = 1;
                    }

                    ExecuteNonQuery(DatabaseConnectionStringNames.Wrappers, "Proc_Insert_Wrapper_Data",
                                        new SqlParameter("@pintLoadsId", loadsId),
                                        new SqlParameter("@pintPageNumber", groupCounter),
                                        new SqlParameter("@pintRecordNumber", groupLineCounter),
                                        new SqlParameter("@pvchrWrapperData", line.Trim()));

                    groupLineCounter++;
                }
            }

            ExecuteNonQuery(DatabaseConnectionStringNames.Wrappers, "Proc_Update_Loads_Count",
                                    new SqlParameter("@pintLoadsID", loadsId),
                                    new SqlParameter("@pintPageCount", groupCounter),
                                    new SqlParameter("@pflgSuccessfulLoad", true));

            WriteToJobLog(JobLogMessageType.INFO, "Load Information Updated.");

        }
    }
}
