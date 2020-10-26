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

namespace SubBalanceLoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Sub Balance Load";
            JobDescription = "Parses a pipe delimited file into a work/staging database";
            AppConfigSectionName = "SubBalanceLoad";

        }

        public override void ExecuteJob()
        {
            try
            {
                //get the input files that are ready for processing
                string inputDirectory = GetConfigurationKeyValue("InputDirectory");
                List<string> files = Directory.GetFiles($"{inputDirectory}\\", "SubBalance.txt").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.SubBalanceLoad, "dbo.Proc_Select_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                        if (previouslyLoadedFile == null)
                        {
                            //make sure the file is no longer being edited
                            if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                            {
                                WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                CopyAndProcessFile(fileInfo);
                            }
                        }
                        //else
                        //{
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.SubBalanceLoad, "Proc_Insert_Loads_Not_Loaded",
                        //                    new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString()),
                        //                    new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                        //                    new SqlParameter("@pdatLastModified", fileInfo.LastWriteTime),
                        //                    new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                        //                    new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                        //                    new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString()));
                        //}
                    }

                    //    }
                    }
                }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void CopyAndProcessFile(FileInfo fileInfo)
        {
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName, true);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.SubBalanceLoad, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("PBSServerInstance")),
                                                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("PBSDatabaseName")),
                                                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("PBSUserName")),
                                                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("PBSPassword")),
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.SubBalanceLoad, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            Int32 lineNumber = 0;

            foreach (string line in fileContents)
            {
                if (line != null && line.Trim().Length > 0)
                {
                    List<string> lineSegments = line.Split('|').ToList();

                    ExecuteNonQuery(DatabaseConnectionStringNames.SubBalanceLoad, "dbo.Proc_Insert_SubsBalance",
                                                   new SqlParameter("@loadsId", loadsId),
                                                   new SqlParameter("@subscriptionId", FormatNumber(lineSegments[0].ToString())),
                                                   new SqlParameter("@balanceDate", FormatDateTime(lineSegments[1].ToString())),
                                                   new SqlParameter("@balance", FormatNumber(lineSegments[2].ToString())),
                                                   new SqlParameter("@unallocated", FormatNumber(lineSegments[3].ToString())),
                                                   new SqlParameter("@discount", FormatNumber(lineSegments[4].ToString())),
                                                   new SqlParameter("@grace", FormatNumber(lineSegments[5].ToString())),
                                                   new SqlParameter("@inGrace", FormatNumber(lineSegments[6].ToString())));

                    lineNumber++;
                   
                }

            }

            WriteToJobLog(JobLogMessageType.INFO, $"{lineNumber} records read.");

            ExecuteNonQuery(DatabaseConnectionStringNames.PBSDump, "dbo.Proc_Insert_SubsBalance",
                                        new SqlParameter("@pvchrServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                        new SqlParameter("@pvchrDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")),
                                        new SqlParameter("@pintLoadsID", loadsId));

            WriteToJobLog(JobLogMessageType.INFO, "PBSDump table successfully repopulated");

            ExecuteNonQuery(DatabaseConnectionStringNames.SubBalanceLoad, "Proc_Delete_SubsBalance",
                                        new SqlParameter("@pintLoadsID", loadsId));

            ExecuteNonQuery(DatabaseConnectionStringNames.SubBalanceLoad, "Proc_Update_Loads_Successful",
                                        new SqlParameter("@pintLoadsID", loadsId));


        }

    }
}
