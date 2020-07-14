using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PressRoomLoad
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "press.*").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.PressRoom, "dbo.Proc_Select_Loads_If_Processed",
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
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.PressRoom, "Proc_Insert_Loads_Not_Loaded",
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

        private void CopyAndProcessFile(FileInfo fileInfo)
        {
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName, true);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PressRoom, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.PressRoom, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            DateTime? publishDate = null;
            DateTime? runDate = null; 
            string edition = null;
            string deliverySelection = null;
            Int32 subtotal = 0;
            List<string> warnings = new List<string>();
            bool inWarningMessage = false;

            foreach (string line in fileContents)
            {

                if (line != null && line.Trim().Length > 0)
                {
                    if (line.StartsWith("PUBLISHING DATE:"))
                        publishDate = Convert.ToDateTime(line.Replace("PUBLISHING DATE:", "").Trim());
                    else if (line.StartsWith("EDITION:"))
                        edition = line.Replace("EDITION:", "").Trim();
                    else if (line.StartsWith("DELIVERY SELECTION:"))
                        deliverySelection = line.Replace("DELIVERY SELECTION:", "").Trim();
                    else if (line.Trim().StartsWith("PRINT SUBTOTAL:"))
                        subtotal = Convert.ToInt32(line.Replace("PRINT SUBTOTAL:", "").Replace(",", "").Trim());
                    else if (line.Trim().StartsWith("WARNING:"))
                    {
                        warnings.Add(line.Trim().Replace("WARNING:", "").Trim());
                        inWarningMessage = true;
                    }
                    else if (inWarningMessage)
                    {
                        warnings.Add(line.Trim());
                    }

                }
                else if (line.StartsWith("\f") && publishDate != null)
                {
                    //form feeds mark the end of a group, so save the values to the database
                    result = ExecuteSQL(DatabaseConnectionStringNames.PressRoom, "dbo.Proc_Insert_Delivery_Selection_Totals",
                                                         new SqlParameter("@pintLoadsID", loadsId),
                                                         new SqlParameter("@psdatRunDate", publishDate),
                                                         new SqlParameter("@pvchrEdition", edition),
                                                         new SqlParameter("@pvchrDeliverySelection", deliverySelection),
                                                         new SqlParameter("@pintProductTotal", subtotal)).FirstOrDefault();

                    Int32 deliverySelectionId = Int32.Parse(result["delivery_selection_totals_id"].ToString());

                    if (warnings.Count() > 0)
                    {
                        int warningCount = 1;
                        foreach (string warning in warnings)
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.PressRoom, "dbo.Proc_Insert_Warnings",
                                                    new SqlParameter("@pintDeliverySelectionTotalsID", deliverySelectionId),
                                                    new SqlParameter("@pintSequence", warningCount),
                                                    new SqlParameter("@pvchrWarning", warning));

                            warningCount++;
                        }
                    }

                    runDate = publishDate;

                    //clear all variables
                    inWarningMessage = false;
                    publishDate = null;
                    edition = null;
                    deliverySelection = null;
                    subtotal = 0;
                    warnings = new List<string>();
                }
            }

            ExecuteNonQuery(DatabaseConnectionStringNames.PressRoom, "dbo.Proc_Update_Loads_Date",
                                                new SqlParameter("@pintLoadsID", loadsId),
                                                new SqlParameter("@sdatRunDate", runDate),
                                                new SqlParameter("@pflgSuccessfullLoad", 1));

            WriteToJobLog(JobLogMessageType.INFO, "Load information updated.");

        }

        public override void SetupJob()
        {
            JobName = "Press Room Load";
            JobDescription = @"Parses a fixed width file with delivery by area totals";
            AppConfigSectionName = "PressRoomLoad";
        }
    }
}
