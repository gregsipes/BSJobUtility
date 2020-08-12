using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace PBSDumpWorkload
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "PBS Dump Workload";
            JobDescription = "";
            AppConfigSectionName = "PBSDumpWorkload";
        }

        public override void ExecuteJob()
        {
            try
            {

                //check for any touch files. If they exist, send email with contents, then delete (is this even needed?)
                if (Directory.Exists(GetConfigurationKeyValue("DumpTouchDirectory")))
                {
                    List<string> processedTouchFiles = new List<string>();
                    List<string> touchFiles = Directory.GetFiles(GetConfigurationKeyValue("DumpTouchDirectory"), GetConfigurationKeyValue("DumpTouchFile")).ToList();

                    foreach (string file in touchFiles)
                    {
                        FileInfo fileInfo = new FileInfo(file);
                        if (fileInfo.Name != "." && fileInfo.Name != ".." && processedTouchFiles.Where(p => p == file).Count() == 0)
                        {
                            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();
                            StringBuilder stringBuilder = new StringBuilder();

                            foreach (string line in fileContents)
                            {
                                if (line.Trim() != "")
                                    stringBuilder.AppendLine(line);
                            }

                            SendMail($"{GetConfigurationKeyValue("DumpTouchDescription")} Started {fileInfo.LastWriteTime.ToString()}", stringBuilder.ToString(), false);                          


                            processedTouchFiles.Add(file);
                            File.Delete(file);
                            WriteToJobLog(JobLogMessageType.INFO, $"Deleted file {file}");
                        }
                    }
                }

                //get the input files that are ready for processing
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "dumpcontrol*.timestamp").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "dbo.Proc_Select_BN_Loads_DumpControl_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                        if (previouslyLoadedFile == null)
                        {
                            //make sure we the file is no longer being edited
                            if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                            {
                                WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                InsertLoad(fileInfo);
                            }
                        }
                        //else
                        //{
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Insert_Loads_Not_Loaded",
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
                LogException(ex);
                throw;
            }
        }

        private void InsertLoad(FileInfo fileInfo)
        {

            WriteToJobLog(JobLogMessageType.INFO, $"Dump control's timestamp = " + fileInfo.LastWriteTime.ToString());


            //create load record
            Int32 loadsId = 0;
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "dbo.Proc_Insert_BN_Loads_DumpControl",
                            new SqlParameter("@pdatTimeStamp", fileInfo.LastWriteTime),
                            new SqlParameter("@pvchrOriginalDir", fileInfo.Directory),
                            new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                            new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                            new SqlParameter("@pvchrUserName", Environment.UserName),
                            new SqlParameter("@pvchrComputerName", Environment.MachineName),
                            new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();

            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads Dump Control ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_DumpControl_BNTimeStamp",
                                 new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                 new SqlParameter("@pvchrBNTimeStamp", fileInfo.LastWriteTime));

            ProcessFile(fileInfo);

        }

        private void ProcessFile(FileInfo fileInfo)
        {
            WriteToJobLog(JobLogMessageType.INFO, $"Reading {fileInfo.Name}");

            if (fileInfo.Length > 0)
            {
                List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

                foreach (string line in fileContents)
                {
                    List<string> segments = line.Split('|').ToList();


                }
            }
        }
    }
}
