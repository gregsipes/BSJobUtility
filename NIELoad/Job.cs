using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace NIELoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "NIE Load";
            JobDescription = "";
            AppConfigSectionName = "NIELoad";
        }

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "nietrack.txt").ToList();

                files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "niedon.txt").ToList());
                files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "nielbx.txt").ToList());
                files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "niecocr.txt").ToList());

                List<string> processedFiles = new List<string>();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.NIEDonations, "dbo.Proc_Select_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                        if (previouslyLoadedFile == null)
                        {
                            //make sure we the file is no longer being edited
                            if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                            {
                                WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                CopyAndProcessFile(fileInfo);
                            }
                        }
                        //else
                        //{
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.NIEDonations, "Proc_Insert_Loads_Not_Loaded",
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
    }
}
