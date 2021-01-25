using BSGlobals;
using BSJobBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace ArchiveAutoRenewals
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Archive Auto Renewals";
            JobDescription = "Archives invoice/renewal PDF's in an effort to keep the folders from growing too large.";
            AppConfigSectionName = "ArchiveAutoRenewals";
        }

        public override void ExecuteJob()
        {
            try
            {
                ArchiveCircFiles();
                ArchiveSFTPFiles();
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void ArchiveSFTPFiles()
        {
            WriteToJobLog(JobLogMessageType.INFO, "Creating SFTP session");

            SFTP sFTP = new SFTP(GetConfigurationKeyValue("HostName"), GetConfigurationKeyValue("UserName"), GetConfigurationKeyValue("Password"));
            sFTP.OpenSession(GetConfigurationKeyValue("FingerPrint"), GetConfigurationKeyValue("KeyFilePath"), GetConfigurationKeyValue("KeyFilePassword"));

            WriteToJobLog(JobLogMessageType.INFO, $"Retrieving files from {GetConfigurationKeyValue("SFTPPath")}");

            List<string> remoteFiles = sFTP.GetFiles(GetConfigurationKeyValue("SFTPPath"), GetConfigurationKeyValue("InputMask"));

            WriteToJobLog(JobLogMessageType.INFO, $"{remoteFiles.Count().ToString()} files retrieved");

            if (remoteFiles != null && remoteFiles.Count() > 0)
            {
                foreach (string remoteFile in remoteFiles)
                {
                    string fileName = remoteFile.Substring(remoteFile.LastIndexOf("/") + 1);

                    if (fileName.Length == 27)
                    {
                        string month = fileName.Substring(8, 2);
                        string day = fileName.Substring(10, 2);
                        string year = fileName.Substring(12, 4);

                        //create year folder if one doesn't already exist
                        string outputPath = Path.Combine(GetConfigurationKeyValue("SFTPPath"), year).Replace("\\", "/");
                        if (!Directory.Exists(outputPath))
                            Directory.CreateDirectory(outputPath);

                        //create month folder if one doesn't already exist
                        outputPath = Path.Combine(outputPath, month).Replace("\\", "/");
                        if (!Directory.Exists(outputPath))
                            Directory.CreateDirectory(outputPath);

                        //create day folder if one doesn't already exist
                        outputPath = Path.Combine(outputPath, day).Replace("\\", "/");
                        if (!Directory.Exists(outputPath))
                            Directory.CreateDirectory(outputPath);

                        //move file to destination folder
                        outputPath = Path.Combine(outputPath, fileName);
                        if (!sFTP.CheckIfFileOrDirectoryExists(outputPath))
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"Moving file {remoteFile} to {outputPath}");
                            sFTP.MoveFile(remoteFile, outputPath);
                        }
                    }
                }
            }

        }

        private void ArchiveCircFiles()
        {

            WriteToJobLog(JobLogMessageType.INFO, $"Retrieving files from {GetConfigurationKeyValue("CircPath")}");

            List<string> existingFiles = Directory.GetFiles(GetConfigurationKeyValue("CircPath"), GetConfigurationKeyValue("InputMask")).ToList();

            WriteToJobLog(JobLogMessageType.INFO, $"{existingFiles.Count().ToString()} files retrieved");

            if (existingFiles != null && existingFiles.Count() > 0)
            {
                foreach (string existingFile in existingFiles)
                {
                    string fileName = existingFile.Substring(existingFile.LastIndexOf("/") + 1);

                    if (fileName.Length == 27)
                    {
                        string month = fileName.Substring(8, 2);
                        string day = fileName.Substring(10, 2);
                        string year = fileName.Substring(12, 4);

                        //create year folder if one doesn't already exist
                        string outputPath = Path.Combine(GetConfigurationKeyValue("CircPath"), year);
                        if (!Directory.Exists(outputPath))
                            Directory.CreateDirectory(outputPath);

                        //create month folder if one doesn't already exist
                        outputPath = Path.Combine(outputPath, month);
                        if (!Directory.Exists(outputPath))
                            Directory.CreateDirectory(outputPath);

                        //create day folder if one doesn't already exist
                        outputPath = Path.Combine(outputPath, day);
                        if (!Directory.Exists(outputPath))
                            Directory.CreateDirectory(outputPath);

                        //move file to destination folder
                        outputPath = Path.Combine(outputPath, fileName);
                        if (!File.Exists(outputPath))
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"Moving file {existingFile} to {outputPath}");
                            File.Move(existingFile, outputPath);
                        }
                    }
                }
            }
        }


    }


}
