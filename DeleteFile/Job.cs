using BSJobBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace DeleteFile
{
    public class Job : JobBase
    {

        public override void SetupJob()
        {
            JobName = "Delete File";
            JobDescription = "Deletes files from directories based on age.";
            AppConfigSectionName = "DeleteFile";
        }

        public override void ExecuteJob()
        {
            try
            {
                //get the files to search for
                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.BSJobUtility, "Proc_Select_Delete_Files").ToList();

                Int32 deletedFileCount = 0;
                Int32 deletedFolderCount = 0;

                foreach (Dictionary<string, object> result in results)
                {

                    DirectoryInfo directoryInfo = new DirectoryInfo(result["Path"].ToString());


                    List<FileInfo> files = directoryInfo.GetFiles(result["FileSearchPattern"].ToString()).ToList();
                    FileInfo mostRecentFile = files.OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                    //delete each file
                    foreach (FileInfo fileInfo in files)
                    {
                        //check to see if the file is older than the number of days to keep the file
                        if (DateTime.Now.AddDays(Convert.ToInt32(result["DaysToKeep"].ToString()) * -1) > fileInfo.LastWriteTime)
                        {
                            //delete the file only if we can either delete all files or this is not the latest version
                            if (Convert.ToBoolean(result["DeleteLatestVersion"].ToString()) == true ||
                                (Convert.ToBoolean(result["DeleteLatestVersion"].ToString()) == false && mostRecentFile.FullName != fileInfo.FullName))
                            {
                                WriteToJobLog(JobLogMessageType.INFO, $"Deleting {fileInfo.FullName}");
                                File.Delete(fileInfo.FullName);
                                deletedFileCount++;
                            }
                        }
                    }

                    //check subdirectories
                    if (Convert.ToBoolean(result["DeleteEmptySubDirectories"].ToString()) || Convert.ToBoolean(result["DeleteAllSubDirectories"].ToString()))
                        deletedFolderCount += DeleteSubDirectories(directoryInfo.FullName, result);
                }

                if (deletedFileCount == 0 && deletedFolderCount == 0)
                    WriteToJobLog(JobLogMessageType.INFO, "No files and folders to delete");
                else
                {
                    WriteToJobLog(JobLogMessageType.INFO, $"{deletedFileCount} files deleted");
                    WriteToJobLog(JobLogMessageType.INFO, $"{deletedFolderCount} folders deleted");
                }

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private Int32 DeleteSubDirectories(string directoryPath, Dictionary<string, object> result)
        {
            Int32 deletedFolderCount = 0;

            List<string> folders = Directory.GetDirectories(directoryPath).ToList();

            foreach (string folder in folders)
            {
                //if the directory is empty, always delete it. If it is not empty, check to see if the flag is set to delete anyways
                if (!Directory.EnumerateFileSystemEntries(folder).Any())
                {
                    WriteToJobLog(JobLogMessageType.INFO, $"Deleting empty folder {folder}");
                    Directory.Delete(folder);
                    deletedFolderCount++;
                }
                else if (Convert.ToBoolean(result["DeleteAllSubDirectories"].ToString()))
                {
                    //only delete the subdirectory if the newest file is older than the days to keep
                    DirectoryInfo directoryInfo = new DirectoryInfo(folder);
                    FileInfo mostRecentFile = directoryInfo.GetFiles("*", SearchOption.AllDirectories).OrderByDescending(f => f.LastWriteTime).FirstOrDefault();

                    if (DateTime.Now.AddDays(Convert.ToInt32(result["DaysToKeep"].ToString()) * -1) > mostRecentFile.LastWriteTime)
                    {
                        WriteToJobLog(JobLogMessageType.INFO, $"Deleting {folder} with all of its contents recursively");
                        Directory.Delete(folder, true);   //recursively deletes everything in folder
                        deletedFolderCount++;
                    }
                }
            }

            return deletedFolderCount;
        }


    }
}
