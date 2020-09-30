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

namespace QualificationReportLoad
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "combined*.pqr").ToList();

              //  List<string> processedFiles = new List<string>();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Length > 0) //ignore empty files
                        {

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.QualificationReportLoad, "dbo.Proc_Select_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                            if (previouslyLoadedFile == null)
                            {
                                WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                CopyAndProcessFile(fileInfo);
                                // processedFiles.Add(fileInfo.Name);
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
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.QualificationReportLoad, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.QualificationReportLoad, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            string startOfRecordsLine = "----- ----- ----- ---------- --------- ----- --- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- --------";
            bool pastStartLine = false;
            string endOfRecordsLine = "---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ---- ----";
            
            string currentSackNumber = "";
            string currentSackLevel = "";
            string currentSackZip = "";

            DateTime? runDate = null;

            foreach (string line in fileContents)
            {
                if (line != null && line.Trim().Length > 0)
                {
                    if (!pastStartLine)
                    {
                        if (line.Trim() == startOfRecordsLine)
                            pastStartLine = true;
                        else if (line.StartsWith("Mailer Name"))
                        {
                            runDate = DateTime.Parse(line.Substring(60, 2) + "/" + line.Substring(62, 2) + "/" + line.Substring(56, 4));
                            ExecuteNonQuery(DatabaseConnectionStringNames.QualificationReportLoad, "dbo.Proc_Update_Loads_Date",
                                                        new SqlParameter("@pintLoadsID", loadsId),
                                                        new SqlParameter("@sdatRunDate", runDate.Value.ToShortDateString()));
                        }

                    }
                    else
                    {
                        //parse each line until we reach the end
                        if (line.Trim() == endOfRecordsLine)
                            break;
                        else
                        {
                            //set sack number if one is present, if not, use the previous one
                            if (line.Substring(0, 6).Trim() != "")
                                currentSackNumber = line.Substring(0, 6).Trim();

                            //set the sack level if one is present, if not, use the previous one
                            if (line.Substring(6, 6).Trim() != "")
                                currentSackLevel = line.Substring(6, 6).Trim();

                            //set the sack zip if one is present, if not use the previous one
                            if (line.Substring(12, 6).Trim() != "")
                                currentSackZip = line.Substring(12, 6).Trim();

                            //check for the quantity record in one of these fields
                            int quantity = 0;
                            if (line.Substring(49, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(49, 5).Trim());
                            else if (line.Substring(54, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(54, 5).Trim());
                            else if (line.Substring(59, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(59, 5).Trim());
                            else if (line.Substring(64, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(64, 5).Trim());
                            else if (line.Substring(69, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(69, 5).Trim());
                            else if (line.Substring(74, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(74, 5).Trim());
                            else if (line.Substring(79, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(79, 5).Trim());
                            else if (line.Substring(84, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(84, 5).Trim());
                            else if (line.Substring(89, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(89, 5).Trim());
                            else if (line.Substring(94, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(94, 5).Trim());
                            else if (line.Substring(99, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(99, 5).Trim());
                            else if (line.Substring(104, 5).Trim() != "")
                                quantity = Int16.Parse(line.Substring(104, 5).Trim());

                            ExecuteNonQuery(DatabaseConnectionStringNames.QualificationReportLoad, "dbo.Proc_Insert_Pieces",
                                                new SqlParameter("@pintLoadsID", loadsId),
                                                new SqlParameter("@pvchrSackNumber", currentSackNumber),
                                                new SqlParameter("@pvchrSackLevel", currentSackLevel),
                                                new SqlParameter("@pvchrSackZip", currentSackZip),
                                                new SqlParameter("@pintPieces", quantity));

                           
                        }
                    }
                   

                }
            }

            ExecuteNonQuery(DatabaseConnectionStringNames.QualificationReportLoad, "dbo.Proc_Update_Loads_Successful",
                                new SqlParameter("@pintLoadsID", loadsId));

        }

        public override void SetupJob()
        {
            JobName = "Qualification Report Load";
            JobDescription = @"Parses a fixed width file from USPS";
            AppConfigSectionName = "QualificationReportLoad";
        }
    }
}
