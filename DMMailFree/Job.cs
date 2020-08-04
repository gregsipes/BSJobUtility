using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace DMMailFree
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "dmexp**").ToList();

                List<string> processedFiles = new List<string>();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.DMMail, "dbo.Proc_Select_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                        if (previouslyLoadedFile == null)
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                            CopyAndProcessFile(fileInfo);
                            processedFiles.Add(fileInfo.Name);
                        }
                        //else
                        //{
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Loads_Not_Loaded",
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
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.DMMail, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")),
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.DMMail, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            Int32 lineCounter = 0;
            DateTime? publishDate = null;


            if (fileContents.Count > 0)
            {
                foreach (string line in fileContents)
                {
                    lineCounter++;

                    List<string> lineSegments = line.Split('|').ToList();

                    if (lineSegments[0] == "D1")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.DMMail, "Proc_Insert_DMMAILData",
                                    new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@RecordType", lineSegments[0]),
                                      new SqlParameter("@ProductId", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                      new SqlParameter("@EditionId", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                      new SqlParameter("@PublishDate", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                      new SqlParameter("@TruckId", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                      new SqlParameter("@TruckName", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                      new SqlParameter("@RelayTruckId", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                      new SqlParameter("@RelayTruckName", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                      new SqlParameter("@DropOrder", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                      new SqlParameter("@DistrictManagerId", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                                      new SqlParameter("@BillToId", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                                      new SqlParameter("@RouteId", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                      new SqlParameter("@SubscriptionId", lineSegments[12].ToString() == "" ? (object)DBNull.Value : lineSegments[12].ToString()),
                                      new SqlParameter("@SubName", lineSegments[13].ToString() == "" ? (object)DBNull.Value : lineSegments[13].ToString()),
                                      new SqlParameter("@Address1", lineSegments[14].ToString() == "" ? (object)DBNull.Value : lineSegments[14].ToString()),
                                      new SqlParameter("@Address2", lineSegments[15].ToString() == "" ? (object)DBNull.Value : lineSegments[15].ToString()),
                                      new SqlParameter("@Address3", lineSegments[16].ToString() == "" ? (object)DBNull.Value : lineSegments[16].ToString()),
                                      new SqlParameter("@Address4", lineSegments[17].ToString() == "" ? (object)DBNull.Value : lineSegments[17].ToString()),
                                      new SqlParameter("@Address5", lineSegments[18].ToString() == "" ? (object)DBNull.Value : lineSegments[18].ToString()),
                                      new SqlParameter("@Phone", lineSegments[19].ToString() == "" ? (object)DBNull.Value : lineSegments[19].ToString()),
                                      new SqlParameter("@DeliveryScheduleId", lineSegments[20].ToString() == "" ? (object)DBNull.Value : lineSegments[20].ToString()),
                                      new SqlParameter("@BillingMethod", lineSegments[21].ToString() == "" ? (object)DBNull.Value : lineSegments[21].ToString()),
                                      new SqlParameter("@TransTypeId", lineSegments[22].ToString() == "" ? (object)DBNull.Value : lineSegments[22].ToString()),
                                      new SqlParameter("@SourceCode", lineSegments[23].ToString() == "" ? (object)DBNull.Value : lineSegments[23].ToString()),
                                      new SqlParameter("@SubSourceCode", lineSegments[24].ToString() == "" ? (object)DBNull.Value : lineSegments[24].ToString()),
                                      new SqlParameter("@ReasonCode", lineSegments[25].ToString() == "" ? (object)DBNull.Value : lineSegments[25].ToString()),
                                      new SqlParameter("@ComplaintTime", lineSegments[26].ToString() == "" ? (object)DBNull.Value : lineSegments[26].ToString()),
                                      new SqlParameter("@DeliveryComplaint", lineSegments[27].ToString() == "" ? (object)DBNull.Value : lineSegments[27].ToString()),
                                      new SqlParameter("@DaysAdjust", lineSegments[28].ToString() == "" ? (object)DBNull.Value : lineSegments[28].ToString()),
                                      new SqlParameter("@EntryDate", lineSegments[29].ToString() == "" ? (object)DBNull.Value : lineSegments[29].ToString()),
                                      new SqlParameter("@TransDate", lineSegments[30].ToString() == "" ? (object)DBNull.Value : lineSegments[30].ToString()),
                                      new SqlParameter("@Text1", lineSegments[31].ToString() == "" ? (object)DBNull.Value : lineSegments[31].ToString()),
                                      new SqlParameter("@Text2", lineSegments[32].ToString() == "" ? (object)DBNull.Value : lineSegments[32].ToString()),
                                      new SqlParameter("@Text3", lineSegments[33].ToString() == "" ? (object)DBNull.Value : lineSegments[33].ToString()),
                                      new SqlParameter("@Text4", lineSegments[34].ToString() == "" ? (object)DBNull.Value : lineSegments[34].ToString()),
                                      new SqlParameter("@Text5 ", lineSegments[35].ToString() == "" ? (object)DBNull.Value : lineSegments[35].ToString()),
                                      new SqlParameter("@Copies", lineSegments[36].ToString() == "" ? (object)DBNull.Value : lineSegments[36].ToString()),
                                      new SqlParameter("@ExpireDate", lineSegments[37].ToString() == "" ? (object)DBNull.Value : lineSegments[37].ToString()),
                                      new SqlParameter("@RestartDate", lineSegments[38].ToString() == "" ? (object)DBNull.Value : lineSegments[38].ToString()),
                                      new SqlParameter("@CreateUser", lineSegments[39].ToString() == "" ? (object)DBNull.Value : lineSegments[39].ToString()),
                                      new SqlParameter("@DeliveryPlacementID", lineSegments[40].ToString() == "" ? (object)DBNull.Value : lineSegments[40].ToString()),
                                      new SqlParameter("@Description", lineSegments[41].ToString() == "" ? (object)DBNull.Value : lineSegments[41].ToString()),
                                      new SqlParameter("@PublicationName", lineSegments[42].ToString() == "" ? (object)DBNull.Value : lineSegments[42].ToString()),
                                      new SqlParameter("@LastField", lineSegments[43].ToString() == "" ? (object)DBNull.Value : lineSegments[43].ToString()));

                        publishDate = DateTime.Parse(lineSegments[3].ToString());
                    }

                }
            }


            WriteToJobLog(JobLogMessageType.INFO, $"{lineCounter} records read for publishing date {(publishDate.HasValue ? publishDate.Value.ToShortDateString() : (object)DBNull.Value)}.");

            ExecuteNonQuery(DatabaseConnectionStringNames.DMMail, "Proc_Insert_Editions",
                                new SqlParameter("@pintLoadsID", loadsId));

            ExecuteNonQuery(DatabaseConnectionStringNames.DMMail, "Proc_Insert_Editions_No_AMPM",
                                new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, $"Editions associated with loads_id {loadsId} saved.");

            ExecuteNonQuery(DatabaseConnectionStringNames.DMMail, "Proc_Insert_Products",
                    new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, $"Products associated with loads_id {loadsId} saved.");

            ExecuteNonQuery(DatabaseConnectionStringNames.DMMail, "Proc_Insert_Loads_Latest",
                            new SqlParameter("@pintLoadsID", loadsId),
                            new SqlParameter("@psdatRunDate", publishDate.HasValue ? publishDate.Value.ToShortDateString() : (object)DBNull.Value),
                            new SqlParameter("@pintRecordCount", lineCounter),
                            new SqlParameter("@pflgSuccessful", true));
            WriteToJobLog(JobLogMessageType.INFO, $"Load information updated.");

        }

        public override void SetupJob()
        {
            JobName = "DMMailFree";
            JobDescription = @"Parses a pipe delimited file containing detailed manifest records - free version";
            AppConfigSectionName = "DMMailFree";
        }

    }
}
