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
                //get print types
                List<Dictionary<string, object>> printTypes = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Select_Print_Types_For_Load").ToList();
                //get carrier identifiers
                List<Dictionary<string, object>> carrierIdentifiers = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Select_Carrier_Identifiers",
                                                                                       new SqlParameter("@pvchrCarrierIdentifier", null));
                //get carrier exceptions
                List<Dictionary<string, object>> carrierExceptions = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Select_Carrier_Exceptions",
                                                                                new SqlParameter("@pvchrCarrier", null));
                //get total identifiers
                List<Dictionary<string, object>> totalIdentifiers = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.PROC_SELECT_TOTAL_IDENTIFIERS",
                                                                                new SqlParameter("@pvchrCarrier", null));

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
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "Proc_Insert_Loads_Not_Loaded",
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

        private void CopyAndProcessFile(FileInfo fileInfo, List<Dictionary<string, object>> printTypes)
        {
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));


            List<Dictionary<string, object>> amountLabels = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Amount_Due_Labels",
                                                                            new SqlParameter("@pvchrAmountDueLabel", null),
                                                                            new SqlParameter("@pflgActiveOnly", 1)).ToList();

            string workingFilePath = GetConfigurationKeyValue("WorkDirectory1") + "carrinv_" + DateTime.Now.ToString("yyMMddhhmmss") + "_" + fileInfo.Name;
            //create a working copy of the file
            File.Copy(fileInfo.FullName, workingFilePath);

            FileInfo workingFileInfo = new FileInfo(workingFilePath);


            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(workingFileInfo.FullName).ToList();

            string printDate = "";
            DateTime? invoiceDate = null;
            string invoiceNumber = "";
            decimal? balanceDue = null;
            string account = "";
            string route = "";
            string district = "";
            string truck = "";
            Int32? sequence = null;
            string pageNumber = "";

            Int32 pageLineNumber = 0;


            decimal printInvoiceChargeTotal = 0;
            decimal printInvoiceCreditTotal = 0;
            bool printInvoiceCharge = false;
            bool printInvoiceCredit = false;
            bool printReturnSheet = false;
            bool retailDraw = false;

            bool checkIdentifiers = false;
            bool checkRouteSuffix = false;

            Dictionary<string, object> printType = null;

            foreach (string line in fileContents)
            {

                if (line != null && line.Trim().Length > 0)
                {

                    if (line.StartsWith("PRINT DATE:"))
                        printDate = line.Replace("PRINT DATE:", "").Trim(); //this is the last statement that we will process in the file
                    else if (line.Contains("PAGE:"))
                        pageNumber = line.Replace("PAGE:", "").Trim();
                    else if (line.Contains("BALANCE DUE:"))
                        balanceDue = Convert.ToDecimal(line.Substring(line.IndexOf("BALANCE DUE:")).Trim().Replace(",", ""));
                    else if (line.Contains("ACCOUNT     :"))
                        account = line.Substring(line.IndexOf("ACCOUNT     :")).Trim();
                    else if (line.Contains("ROUTE       :"))
                        route = line.Substring(line.IndexOf("ROUTE       :")).Trim();
                    else if (line.Contains("DISTRICT    :"))
                        district = line.Substring(line.IndexOf("DISTRICT    :")).Trim();
                    else if (line.Contains("TRUCK       :"))
                        truck = line.Substring(line.IndexOf("TRUCK       :")).Trim();
                    else if (line.Contains("SEQUENCE    :"))
                        sequence = Convert.ToInt32(line.Substring(line.IndexOf("SEQUENCE    :")).Trim());
                    else if (line.Contains("\f"))
                    {
                        //save values and reset flags
                        //ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices,
                        //       new SqlParameter("@loads_id", loadsId),
                        //        new SqlParameter("@header_page_number", ),
                        //        new SqlParameter("@group_page_number", ),
                        //        new SqlParameter("@print_type", ),
                        //        new SqlParameter("@line_2_identifier", ),
                        //        new SqlParameter("@invoice_number", ),
                        //        new SqlParameter("@invoice_date", ),
                        //        new SqlParameter("@billing_terms", ),
                        //        new SqlParameter("@balance_due", ),
                        //        new SqlParameter("@carrier", ),
                        //        new SqlParameter("@route", route),
                        //        new SqlParameter("@district", ),
                        //        new SqlParameter("@depot", ),
                        //        new SqlParameter("@truck", truck),
                        //        new SqlParameter("@sequence", sequence),
                        //        new SqlParameter("@name_address_1", ),
                        //        new SqlParameter("@name_address_2", ),
                        //        new SqlParameter("@name_address_3", ),
                        //        new SqlParameter("@name_address_4", ),
                        //        new SqlParameter("@barcode", ),
                        //        new SqlParameter("@barcode_readable", ),
                        //        new SqlParameter("@retail_draw_flag", ),
                        //        new SqlParameter("@print_returns_sheet_flag", ),
                        //        new SqlParameter("@print_invoice_charge_flag", ),
                        //        new SqlParameter("@print_invoice_charge", ),
                        //        new SqlParameter("@print_invoice_credit_flag", ),
                        //        new SqlParameter("@print_invoice_credit", ),
                        //        new SqlParameter("@corporate_spreadsheet_amount", ),
                        //        new SqlParameter("@corporate_spreadsheet_retail_daily_draw", ),
                        //        new SqlParameter("@corporate_spreadsheet_retail_daily_draw_charges", ),
                        //        new SqlParameter("@corporate_spreadsheet_retail_sunday_draw", ),
                        //        new SqlParameter("@corporate_spreadsheet_retail_sunday_draw_charges", ),
                        //        new SqlParameter("@corporate_spreadsheet_daily_returns", ),
                        //        new SqlParameter("@corporate_spreadsheet_daily_return_credits", ),
                        //        new SqlParameter("@corporate_spreadsheet_sunday_returns", ),
                        //        new SqlParameter("@corporate_spreadsheet_sunday_return_credits", ),
                        //        new SqlParameter("@corporate_spreadsheet_discount_credits", ),
                        //        new SqlParameter("@corporate_spreadsheet_daily_draw_adj_draw", ),
                        //        new SqlParameter("@corporate_spreadsheet_daily_draw_adj_charges", ),
                        //        new SqlParameter("@corporate_spreadsheet_sunday_draw_adj_draw", ),
                        //        new SqlParameter("@corporate_spreadsheet_sunday_draw_adj_charges", ));

                        pageLineNumber = 0;
                    }

                    pageLineNumber++;


                    switch (pageLineNumber)
                    {
                        case 1:
                            pageNumber = line.Replace("PAGE:", "").Trim();

                            if (pageNumber == "1")
                            {
                                printInvoiceChargeTotal = 0;
                                printInvoiceCreditTotal = 0;
                                printInvoiceCharge = false;
                                printInvoiceCredit = false;
                                printReturnSheet = false;
                                retailDraw = false;
                            }

                            break;
                        case 2:
                            List<string> lineSegments = line.Trim().Split(' ').ToList();
                            printType = printTypes.Where(p => p["line_2_identifier_1"].ToString() == lineSegments[0]).FirstOrDefault(); 

                            if (printType == null)
                                printType = printTypes.Where(p => p["line_2_identifier_2"].ToString() == lineSegments[0]).FirstOrDefault();

                            if (printType != null)
                            {
                                checkIdentifiers = !bool.Parse(printType["do_not_check_carrier_identifiers_flag"].ToString());
                                checkRouteSuffix = bool.Parse(printType["check_route_suffix_flag"].ToString());

                                invoiceDate = DateTime.Parse(line.Replace(printType["line_2_identifier_1"].ToString(), "").Replace(printType["line_2_identifier_2"].ToString(), "").Replace("DATE:", "").Trim());
                            }

                            break;
                        case 3:
                            if (printType["line_2_identifier_1"].ToString() == "Invoice")
                                invoiceNumber = line.Substring(line.IndexOf("INVOICE NO. :")).Replace("INVOICE NO. :", "").Trim();

                            break;

                        case 4:

                            break;
                    }


                         

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

                    //                ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_Control_Process",
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

                    //                ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_Draw_Rate",
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

            //todo: delete file from working directory
            File.Delete(workingFilePath);

            //ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Update_Loads",
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
