using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace PayByScanLoad711
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Pay By Scan Load - 711";
            JobDescription = "Parses a tilda delimited file for 7-11";
            AppConfigSectionName = "PayByScanLoad711";
        }

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "*").ToList();

                //load configuration from configuration specific tables
                Dictionary<string, object> loadFormat = ExecuteSQL(DatabaseConnectionStringNames.PayByScan, "dbo.Proc_Select_Load_Formats",
                                                                                        new SqlParameter("@pvchrLoadFormat", "Seven_Eleven")).FirstOrDefault();
                Dictionary<string, object> fields = ExecuteSQL(DatabaseConnectionStringNames.PayByScan, "dbo.Proc_Select_Load_Formats_Columns",
                                                                                        new SqlParameter("@pvchrLoadFormat", "Seven_Eleven")).FirstOrDefault();

                //iterate and process files
                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Length > 0) //ignore empty files
                        {

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.PayByScan, "dbo.Proc_Select_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                            if (previouslyLoadedFile == null)
                            {
                                //make sure we the file is no longer being edited
                                if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                                {
                                    WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                    CopyAndProcessFile(fileInfo, loadFormat);
                                }
                                else
                                    WriteToJobLog(JobLogMessageType.INFO, "There's a chance the file is still getting updated, so we'll pick it up next run");

                            }
                            //else
                            //{
                            //    ExecuteNonQuery(DatabaseConnectionStringNames.PayByScan, "Proc_Insert_Loads_Not_Loaded",
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

        private void CopyAndProcessFile(FileInfo fileInfo, Dictionary<string, object> loadFormat)
        {
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name.Replace(".csv", "") + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".csv";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName, true);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);


            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PayByScan, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDirectory", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");


            ExecuteNonQuery(DatabaseConnectionStringNames.PayByScan, "Proc_Update_Loads_Backup",
                    new SqlParameter("@pintLoadsID", loadsId),
                    new SqlParameter("@pvchrBackupDirectory", GetConfigurationKeyValue("BackupDirectory")),
                    new SqlParameter("@pvchrBackupFile", fileInfo.Name.Replace(".csv", "") + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".csv"));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();
           // Regex csvParser = new Regex(",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))");

            Int32 lineNumber = 1;

            foreach (string line in fileContents)
            {
                if (line.Count(l => l == ',') == 9) //make sure the line is formatted as expected
                {
                    if (!(lineNumber == 1 && Convert.ToBoolean(loadFormat["column_names_in_first_record_flag"]))) //make sure the first row isn't column headers
                    {
                        List<string> lineSegments = line.Split(Convert.ToChar(loadFormat["field_delimiter"].ToString())).ToList();
                      //  List<string> lineSegments = csvParser.Split(line).ToList();

                        ExecuteNonQuery(DatabaseConnectionStringNames.PayByScan, "Proc_Insert_Seven_Eleven_Work",
                                new SqlParameter("@loads_id", loadsId),
                             new SqlParameter("@record_number", lineNumber),
                             new SqlParameter("@error_message", DBNull.Value),
                             new SqlParameter("@retailer_id", FormatString(lineSegments[0].ToString().Replace("\"", ""))),
                             new SqlParameter("@upc_code", FormatNumber(lineSegments[1].ToString().Replace("\"", ""))),
                             new SqlParameter("@store_number", FormatString(lineSegments[2].ToString().Replace("\"", ""))),
                             new SqlParameter("@item_description", FormatString(lineSegments[3].ToString().Replace("\"", ""))),
                             new SqlParameter("@post_date", FormatString(lineSegments[4].ToString().Replace("\"", ""))),
                             new SqlParameter("@quantity_sold", FormatString(lineSegments[5].ToString().Replace("\"", ""))),
                             new SqlParameter("@unit_cost", FormatString(lineSegments[6].ToString().Replace("\"", ""))),
                             new SqlParameter("@extended_cost", FormatString(lineSegments[7].ToString().Replace("\"", ""))),
                             new SqlParameter("@sale_date", FormatString(lineSegments[8].ToString().Replace("\"", ""))),
                             new SqlParameter("@wholesaler_route", FormatString(lineSegments[9].ToString().Replace("\"", ""))));

                    }
                }

                lineNumber++;
            }

            WriteToJobLog(JobLogMessageType.INFO, $"{lineNumber - 2} total records read.");

            ExecuteNonQuery(DatabaseConnectionStringNames.PayByScan, "Proc_Insert_Seven_Eleven",
                                new SqlParameter("@pintLoadsID", loadsId),
                                new SqlParameter("@pvchrPBSInvoiceExportServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                new SqlParameter("@pvchrPBSInvoiceExportDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")));


            ExecuteNonQuery(DatabaseConnectionStringNames.PayByScan, "Proc_Update_Loads_Data_Record_Count",
                    new SqlParameter("@pintLoadsID", loadsId),
                    new SqlParameter("@pintDataRecordCount", lineNumber - 2),
                    new SqlParameter("@pflgSuccessfulLoad", true));

            WriteToJobLog(JobLogMessageType.INFO, $"Load information updated.");

            ExecuteNonQuery(DatabaseConnectionStringNames.PayByScan, "Proc_Delete_Seven_Eleven_Work",
                    new SqlParameter("@pintLoadsID", loadsId));

            WriteToJobLog(JobLogMessageType.INFO, $"Work records deleted.");

        }
    }
}
