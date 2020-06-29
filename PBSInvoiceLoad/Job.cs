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

                              //  CopyAndProcessFile(fileInfo);

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

            List<Dictionary<string, object>> carrierIdentifiers = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "PROC_SELECT_CARRIER_IDENTIFIERS",
                                                                            new SqlParameter("@pvchrCarrierIdentifier", null)).ToList();

            List<Dictionary<string, object>> carrierExceptions = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "PROC_SELECT_CARRIER_EXCEPTIONS",
                                                                new SqlParameter("@pvchrCarrier", null)).ToList();


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
            string billingTerms = "";
            string nameAddress1 = "";
            string nameAddress2 = "";
            string carrier = "";
            string carrierException = "";
            string pageNumber = "";
            string printTypeIdentifier = "";
            string exceptionPrintType = "";
            string company;

            Int32 pageLineNumber = 0;


            decimal printInvoiceChargeTotal = 0;
            decimal printInvoiceCreditTotal = 0;
            bool printInvoiceCharge = false;
            bool printInvoiceCredit = false;
            bool printReturnSheet = false;
            bool retailDraw = false;
            bool hasCarrierExceptions = false;

            bool checkIdentifiers = false;
            bool checkRouteSuffix = false;

          //  Dictionary<string, object> printType = null;

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

                            // Check for value on line 2 to determine the print type & set flags.
                            // Carrier (including carrier exceptions array) & route may also determine print type.
                            //List<string> lineSegments = line.Trim().Split(' ').ToList();
                            //printType = printTypes.Where(p => p["line_2_identifier_1"].ToString() == lineSegments[0]).FirstOrDefault();

                            //if (printType == null)
                            //{
                            //    printType = printTypes.Where(p => p["line_2_identifier_2"].ToString() == lineSegments[0]).FirstOrDefault();
                            //    printTypeIdentifier = printType["line_2_identifier_2"].ToString();
                            //}
                            //else
                            //    printTypeIdentifier = printType["line_2_identifier_1"].ToString();




                            //if (printType != null)
                            //{
                            //    checkIdentifiers = !bool.Parse(printType["do_not_check_carrier_identifiers_flag"].ToString());
                            //    checkRouteSuffix = bool.Parse(printType["check_route_suffix_flag"].ToString());

                            //    invoiceDate = DateTime.Parse(line.Replace(printType["line_2_identifier_1"].ToString(), "").Replace(printType["line_2_identifier_2"].ToString(), "").Replace("DATE:", "").Trim());
                            //}

                            foreach (Dictionary<string, object> printType in printTypes)
                            {
                                if (line.Contains(printType["line_2_identifier_1"].ToString())) {
                                    printTypeIdentifier = printType["line_2_identifier_1"].ToString();

                                    invoiceDate = DateTime.Parse(line.Replace(printType["line_2_identifier_1"].ToString(), "").Replace("DATE:", "").Trim());
                                    checkIdentifiers = !bool.Parse(printType["do_not_check_carrier_identifiers_flag"].ToString());
                                    checkRouteSuffix = bool.Parse(printType["check_route_suffix_flag"].ToString());

                                    break;
                                }
                             }

                            foreach (Dictionary<string, object> printType in printTypes)
                            {
                                if (line.Contains(printType["line_2_identifier_2"].ToString()))
                                {
                                    printTypeIdentifier = printType["line_2_identifier_2"].ToString();

                                    invoiceDate = DateTime.Parse(line.Replace(printType["line_2_identifier_2"].ToString(), "").Replace("DATE:", "").Trim());
                                    checkIdentifiers = !bool.Parse(printType["do_not_check_carrier_identifiers_flag"].ToString());
                                    checkRouteSuffix = bool.Parse(printType["check_route_suffix_flag"].ToString());

                                    break;
                                }
                            }

                            break;
                        case 3:
                            if (printTypeIdentifier == "Invoice")
                                invoiceNumber = line.Substring(line.IndexOf("INVOICE NO. :")).Replace("INVOICE NO. :", "").Trim();

                            break;
                        case 6:
                            billingTerms = line.Replace("BILLING TERMS:", "").Trim();
                            break;
                        case 8:
                            //check for any possible label values (BALANCE DUE, TOTAL BALANCE DUE, TOTAL OUTSTANDING BALANCE) 
                            foreach (Dictionary<string, object> amountLabel in amountLabels)
                            {
                                if (line.Contains(amountLabel["amount_due_label"].ToString()))
                                {
                                    balanceDue = decimal.Parse(line.Replace(amountLabel["amount_due_label"].ToString(), "").Replace(",", "").Trim());
                                    break;
                                }
                            }
                            break;
                        case 10:
                            nameAddress1 = line.Trim().Substring(0, 40).Trim();
                            carrier = line.Trim().Replace(nameAddress1, "").Replace("ACCOUNT     :", "").Trim();

                            //Check if carrier/print type combinations is in carrier exceptions array.
                            if (hasCarrierExceptions && printTypeIdentifier != "")
                            {
                                foreach (Dictionary<string, object> exception in carrierExceptions)
                                {
                                    if (exception["carrier"].ToString() == carrier)
                                    {
                                        if (exception["print_type_1"].ToString() != "" && printTypeIdentifier == exception["line_2_identifier_1"].ToString())  //print_type_1 is never empty, maybe this is something can happen in the UI?
                                        {
                                            checkIdentifiers = false;
                                            checkRouteSuffix = false;
                                            exceptionPrintType = exception["print_type_1"].ToString();
                                        }
                                    }
                                    else if (exception["print_type_2"].ToString() != "" && printTypeIdentifier == exception["line_2_identifier_2"].ToString())  //print_type_2 is never empty, maybe this is something can happen in the UI?
                                    {
                                        checkIdentifiers = false;
                                        checkRouteSuffix = false;
                                        exceptionPrintType = exception["print_type_2"].ToString();
                                    }
                                    else if (exception["print_type_3"].ToString() != "" && printTypeIdentifier == exception["line_2_identifier_3"].ToString())  //this condition is never currently hit
                                    {
                                        checkIdentifiers = false;
                                        checkRouteSuffix = false;
                                        exceptionPrintType = exception["print_type_3"].ToString();
                                    }
                                }
                            }

                            //Check the carrier to determine print type.
                            if (checkIdentifiers && carrierIdentifiers != null)
                            {
                                foreach (Dictionary<string, object> carrierIdentifier in carrierIdentifiers)
                                {
                                    if (carrier.StartsWith(carrierIdentifier["carrier"].ToString()))
                                    {
                                        exceptionPrintType = carrierIdentifier["print_type"].ToString();
                                        checkRouteSuffix = false;
                                        break;
                                    }
                                }
                            }

                            break;
                        case 11:
                            if (exceptionPrintType == printTypes.Where(p => Boolean.Parse(p["total_flag"].ToString()) == true).Select(p => p["print_type"].ToString()).FirstOrDefault())
                                company = line.Substring(20, 20); //todo, is this correct?
                            else
                            {
                                nameAddress2 = line.Trim().Substring(0, 40).Trim();
                                route = line.Trim().Replace(nameAddress2, "").Trim().Replace("ROUTE       :", "").Trim();
                            }

                            //Check the route to determine print type.
                            if (checkRouteSuffix)
                            {
                                foreach(Dictionary<string, object> printType in printTypes)
                                {
                                   // if (printType["check_route_suffix_flag"])
                                }
                            }


                            break;

                    }




                  
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
