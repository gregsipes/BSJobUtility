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

namespace PrepackInsertLoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Prepack Insert Load";
            JobDescription = "Parses a fixed width file of customer and advertising zones";
            AppConfigSectionName = "PrepackInsertLoad";
        }

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "*").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Length > 0) //ignore empty files
                        {

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.PrepackInsertLoad, "dbo.Proc_Select_Loads_If_Processed",
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
                            //    ExecuteNonQuery(DatabaseConnectionStringNames.PrepackInsertLoad, "Proc_Insert_Loads_Not_Loaded",
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
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PrepackInsertLoad, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.PrepackInsertLoad, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            bool allEditions = false;
            bool columnHeadersFound = false;
            Int32 recordCounter = 0;
            string mixName = "";
            string quantity = "";
            Int32 lineNumber = 1;

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();


            foreach (string line in fileContents)
            {
                if (line.Trim() != "")
                {
                    if (line.Contains("ALL EDITIONS, ")) //this means that we are on a new page
                    {
                       allEditions = true;
                        columnHeadersFound = false;
                        mixName = "";
                    }
                    else

                    if (allEditions)
                    {
                        if (line.StartsWith("PRINTED BY:"))
                        {
                            DateTime mixDate = Convert.ToDateTime(line.Substring(line.IndexOf(",") + 1, 11).Trim());
                            ExecuteNonQuery(DatabaseConnectionStringNames.PrepackInsertLoad, "Proc_Update_Loads_Mix_Date",
                                                        new SqlParameter("@pintLoadsID", loadsId),
                                                        new SqlParameter("@psdatMixDate", mixDate));
                        }
                        else if (line == "RGNL ZONE MIX NAME      ROP PAGES  QUANTITY AD ZONE   QUANTITY CUSTOMER NAME                             VERSION      ID CODE")
                        {
                            columnHeadersFound = true;
                        }
                        else if (columnHeadersFound)
                        {
                            if (line.Substring(10, 13).Trim() != "")
                                mixName = line.Substring(10, 13).Trim();


                            if (mixName == "GRAND TOTAL")
                            {
                                columnHeadersFound = false;
                                mixName = "";
                                continue;
                            }

                            string customer = "";
                            string customerName = "";
                            string adZone = line.Substring(44, 8).Trim();

                            if (adZone == "--------")
                                continue;


                            if (line.Substring(53, 9).Trim() != "---------")
                               quantity= FormatNumber(line.Substring(53, 9).Trim()).ToString();

                            if (line.Length >= 64)
                                customer = line.Substring(63, 8).Trim();

                            if (line.Length >= 72)
                                customerName = line.Substring(72).Trim();


                            if (mixName != "")
                            {
                                if (adZone != "" && adZone != "SUBTOTAL")
                                {
                                    ExecuteNonQuery(DatabaseConnectionStringNames.PrepackInsertLoad, "Proc_Insert_Ad_Zones",
                                                        new SqlParameter("@pintLoadsID", loadsId),
                                                        new SqlParameter("@vchrMixName", mixName),
                                                        new SqlParameter("@pvchrAdZone", adZone),
                                                        new SqlParameter("@pintQuantity", quantity));
                                }

                                if (customer != "" && customerName != "" && customer != "********" && customerName != "MIX CONTINUED ON NEXT PAGE")
                                {
                                    ExecuteNonQuery(DatabaseConnectionStringNames.PrepackInsertLoad, "Proc_Insert_Customers",
                                                        new SqlParameter("@pintLoadsID", loadsId),
                                                        new SqlParameter("@vchrMixName", mixName),
                                                        new SqlParameter("@pvchrCustomer", customer),
                                                        new SqlParameter("@pvchrCustomerName", customerName));
                                }
                            }

                            recordCounter++;
                        }
                    }
                }

                lineNumber++;

            }

            WriteToJobLog(JobLogMessageType.INFO, $"{recordCounter} total records read.");

            ExecuteNonQuery(DatabaseConnectionStringNames.PrepackInsertLoad, "dbo.Proc_Update_Loads_Successful",
                                        new SqlParameter("@pintLoadsID", loadsId));

        }
    }
}
