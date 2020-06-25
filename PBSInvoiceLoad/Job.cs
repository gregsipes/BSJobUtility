using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PBSInvoiceLoad
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invoic*").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Select_Loads_If_Processed",
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
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceTotals, "Proc_Insert_Loads_Not_Loaded",
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
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceTotals, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceTotals, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            //bool inTotalsSection = false;
            //bool inControlTotalsSection = false;
            //bool inDrawSummarySection = false;
            //// bool inGLSection = false;
            //int controlProcessRecordNumber = 0;
            //int drawRecordNumber = 0;
            //// int GLRecordNumber = 0;
            //string billDate = null;
            //string billSource = null;

            foreach (string line in fileContents)
            {

                if (line != null && line.Trim().Length > 0)
                {
                //    if (!inTotalsSection)  //we are only processing the bottom portion of the file
                //    {
                //        if (line.Contains("ACCOUNT     : TOTAL"))
                //            inTotalsSection = true;
                //    }
                //    else
                //    {
                //        if (line.Contains("BILL SOURCE:"))
                //            billSource = line.Substring(0, line.IndexOf("DISTRICT    :")).Replace("BILL SOURCE:", "").Trim();
                //        else if (line.Contains("BILL DATE  :"))
                //            billDate = line.Substring(0, line.IndexOf("TRUCK       :")).Replace("BILL DATE  :", "").Trim();
                //        else if (line.Contains("CONTROL TOTALS"))
                //        {
                //            inControlTotalsSection = true;
                //            inDrawSummarySection = false;
                //            //    inGLSection = false;
                //        }
                //        else if (line.Contains("DRAW SUMMARY"))
                //        {
                //            inControlTotalsSection = false;
                //            inDrawSummarySection = true;
                //            //  inGLSection = false;
                //        }
                //        //else if (line.Contains("GENERAL LEDGER"))
                //        //{
                //        //    inControlTotalsSection = false;
                //        //    inDrawSummarySection = false;
                //        //  //  inGLSection = true;
                //        //}
                //        else if (inControlTotalsSection)
                //        {
                //            decimal controlTotal = 0;
                //            if (decimal.TryParse(line.Substring(30, 15).Trim(), out controlTotal))
                //            {
                //                string description = line.Substring(0, line.IndexOf(".")).Trim();
                //                decimal processTotal = decimal.Parse(line.Substring(46).Trim().Replace(",", ""));

                //                controlProcessRecordNumber++;

                //                ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceTotals, "dbo.Proc_Insert_Control_Process",
                //                                    new SqlParameter("@pintLoadsID", loadsId),
                //                                    new SqlParameter("@pintRecordNumber", controlProcessRecordNumber),
                //                                    new SqlParameter("@pvchrDescription", description),
                //                                    new SqlParameter("@pfltControlTotal", controlTotal),
                //                                    new SqlParameter("@pfltProcessTotal", processTotal));
                //            }
                //        }
                //        else if (inDrawSummarySection)
                //        {
                //            if (line.Contains("@") && line.IndexOf("@") == 42)
                //            {
                //                string description = line.Substring(0, 32).Trim();
                //                Int32 drawTotal = Int32.Parse(line.Substring(0, line.IndexOf("@")).Replace(description, "").Replace(",", "").Trim());
                //                decimal rate = decimal.Parse(FormatNumber(line.Substring(43, 11).Trim().Replace(",", "")).ToString());
                //                decimal total = decimal.Parse(FormatNumber(line.Substring(54).Trim().Replace(",", "")).ToString());

                //                drawRecordNumber++;

                //                ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceTotals, "dbo.Proc_Insert_Draw_Rate",
                //                                    new SqlParameter("@pintLoadsID", loadsId),
                //                                    new SqlParameter("@pintRecordNumber", controlProcessRecordNumber),
                //                                    new SqlParameter("@pvchrDescription", description),
                //                                    new SqlParameter("@pintDrawTotal", drawTotal),
                //                                    new SqlParameter("@pmnyRate", rate),
                //                                    new SqlParameter("@pmnyTotalAmount", total));




                //            }
                //        }
                //        //else if (inGLSection)
                //        //{

                //        //    GLRecordNumber++;
                //        //a new GL record hasn't been created since 2007 

                //        //}
                //    }
                }

            }

            //ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceTotals, "dbo.Proc_Update_Loads",
            //                                    new SqlParameter("@pintLoadsID", loadsId),
            //                                    new SqlParameter("@pvchrBillSource", billSource),
            //                                    new SqlParameter("@pvchrBillDate", billDate));

            WriteToJobLog(JobLogMessageType.INFO, "Load information updated.");

        }

        public override void SetupJob()
        {
            JobName = "PBS Invoices";
            JobDescription = @"";
            AppConfigSectionName = "PBSInvoices";
        }
    }
}
