using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace CopyIfNewer
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "CopyIfNewer";
            JobDescription = "Checks for and copies over the most the most recent version of files, either file by file or an entire directory";
            AppConfigSectionName = "CopyIfNewer";
        }

        public override void ExecuteJob()
        {
            try
            {

              //  WriteToJobLog(JobLogMessageType.INFO, "Executing Proc_Select_CopyIfNewer_Last_Copy_Date_Times");

                Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Newshole, "Proc_Select_CopyIfNewer_Last_Copy_Date_Times",
                                                                new SqlParameter("@pvchrSourceDirectory", GetConfigurationKeyValue("BradburySourceDirectory"))).FirstOrDefault();

                DateTime lastCopyDateTime = Convert.ToDateTime(result["last_copy_date_time"].ToString());


                // List<string> files = Directory.GetFiles(GetConfigurationKeyValue("BradburySourceDirectory"), "*", SearchOption.AllDirectories).Where(f => new FileInfo(f).LastWriteTime.CompareTo(lastCopyDateTime) > 0).ToList();
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("BradburySourceDirectory"), "*", SearchOption.AllDirectories).Take(10).ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        //todo: check for name
                        string parentFolderName = fileInfo.DirectoryName.Substring(fileInfo.DirectoryName.LastIndexOf("\\") + 1);
                        string newFileName = parentFolderName + "_" + fileInfo.Name;

                        //check to see if file already exists and if it is newer
                        if (File.Exists(GetConfigurationKeyValue("BradburyDestinationDirectory") + fileInfo.Name))
                        {
                            FileInfo existingFile = new FileInfo(GetConfigurationKeyValue("BradburyDestinationDirectory") + fileInfo.Name);

                            if (fileInfo.LastWriteTime <= existingFile.LastWriteTime)
                                continue;
                        }

                        //copy the file over and append the source folder to the file name
                        File.Copy(fileInfo.FullName, GetConfigurationKeyValue("BradburyDestinationDirectory") + fileInfo.Name);
                        WriteToJobLog(JobLogMessageType.INFO, $"Copied {fileInfo.FullName} to {GetConfigurationKeyValue("BradburyDestinationDirectory") + fileInfo.Name}");

                    }

                    //update database record
                    WriteToJobLog(JobLogMessageType.INFO, "Executing Proc_Update_CopyIfNewer_Last_Copy_Date_Times");

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
