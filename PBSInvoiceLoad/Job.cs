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
            string billSource = "";
            string nameAddress1 = "";
            string nameAddress2 = "";
            string nameAddress3 = "";
            string nameAddress4 = "";
            DateTime? billDate = null;
            string sortOrder = "";

            string carrier = "";
           // string carrierException = "";
            string pageNumber = "";
            string printTypeIdentifier = "";
            string exceptionPrintType = "";
            string company = "";

            Int32 pageLineNumber = 0;
            Int32 headerLineNumber = 0;
            Int32 bodyLineNumber = 0;
            Int64 invoiceCount = 0;
            Int64 statementCount = 0;
            Int64 totalCount = 0;

            List<string> headerLines = new List<string>();
            List<string> bodyLines = new List<string>();

            decimal printInvoiceChargeTotal = 0;
            decimal printInvoiceCreditTotal = 0;
            //bool printInvoiceCharge = false;
            //bool printInvoiceCredit = false;
            //bool printReturnSheet = false;
            //bool retailDraw = false;
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
                    else if (line.Contains("\f"))
                    {
                        CreateHeaderRecord();
                        

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

                            foreach (Dictionary<string, object> printType in printTypes)
                            {
                                if (line.Contains(printType["line_2_identifier_1"].ToString()))
                                {
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
                                foreach (Dictionary<string, object> printType in printTypes)
                                {
                                    if (Int32.Parse(printType["check_route_suffix_flag"].ToString()) != 0 &
                                        ((Int32.Parse(printType["check_route_alpha_flag"].ToString()) != 0 & route.Substring(route.Length - 2, 1).All(char.IsNumber))
                                                 | (Int32.Parse(printType["check_route_alpha_flag"].ToString()) == 0 & !route.Substring(route.Length - 2, 1).All(char.IsNumber))) &
                                         (printType["line_2_identifier_1"].ToString() == printTypeIdentifier | printType["line_2_identifier_2"].ToString() == printTypeIdentifier))
                                    {
                                        exceptionPrintType = printType["print_type"].ToString();
                                        break;
                                    }
                                }
                            }

                            break;
                        case 12:
                            if (exceptionPrintType == printTypes.Where(p => p["total_flag"].ToString() == "1").Select(p => p["printType"].ToString()).FirstOrDefault())
                                billSource = line.Substring(20, 20).Trim();
                            else
                                nameAddress3 = line.Trim().Substring(0, 40).Trim();

                            break;
                        case 13:
                            if (exceptionPrintType == printTypes.Where(p => p["total_flag"].ToString() == "1").Select(p => p["printType"].ToString()).FirstOrDefault())
                                DateTime.TryParse(line.Substring(20, 20).Trim(), out billDate);
                            else
                                nameAddress4 = line.Trim().Substring(0, 40).Trim();

                            break;
                        case 14:
                            if (exceptionPrintType == printTypes.Where(p => p["total_flag"].ToString() == "1").Select(p => p["printType"].ToString()).FirstOrDefault())
                                sortOrder = line.Substring(20, 20).Trim();

                            break;
                    }

                    if (exceptionPrintType != printTypes.Where(p => p["total_flag"].ToString() == "1").Select(p => p["printType"].ToString()).FirstOrDefault())
                    {
                        if (pageLineNumber >= 12 & pageLineNumber <= 15)
                        {
                            if (line.Contains("DISTRICT    :"))
                                district = line.Substring(line.IndexOf("DISTRICT    :")).Trim();
                            else if (line.Contains("TRUCK       :"))
                                truck = line.Substring(line.IndexOf("TRUCK       :")).Trim();
                            else if (line.Contains("DEPOT       :"))
                                truck = line.Substring(line.IndexOf("DEPOT       :")).Trim();
                            else if (line.Contains("SEQUENCE    :"))
                                sequence = Convert.ToInt32(line.Substring(line.IndexOf("SEQUENCE    :")).Trim());
                        }
                    }

                    if (pageLineNumber > 0 && pageLineNumber <= 16)
                    {
                        headerLineNumber++;
                        headerLines.Add(line);
                    }
                    else if (pageLineNumber > 16)
                    {
                        bodyLineNumber++;
                        bodyLines.Add(line);

                        if (line.Contains("RETAIL DAILY DRAW") | line.Contains("RETAIL SUNDAY DRAW") | line.Contains("CORP STORE DELIVERY CREDIT") |
                            line.Contains("DIRECT BILL DELIVERY CREDIT") | line.Contains("RETURN CREDITS") | line.Contains("USA RETAIL HONOR BOX CHARGE"))
                        {
                            printReturnSheet = true;
                        }

                        if (line.Contains("RETAIL DAILY DRAW") | line.Contains("RETAIL SUNDAY DRAW"))
                            retailDraw = true;

                        if (line.Contains("PRINT INVOICE CHARGE"))
                        {
                            printInvoiceCharge = true;
                            printInvoiceChargeTotal = Convert.ToDecimal(FormatNumber(line.Substring(line.IndexOf("PRINT INVOICE CHARGE")).Replace("PRINT INVOICE CHARGE", "").Trim()));
                        }

                        if (line.Contains("PRINT INVOICE CREDIT"))
                        {
                            printInvoiceCredit = true;
                            printInvoiceCreditTotal = Convert.ToDecimal(FormatNumber(line.Substring(line.IndexOf("PRINT INVOICE CREDIT")).Replace("PRINT INVOICE CREDIT", "").Trim()));
                        }
                    }

                    if (exceptionPrintType != printTypes.Where(p => p["total_flag"].ToString() == "1").Select(p => p["printType"].ToString()).FirstOrDefault())
                    {
                        if (line.Contains("Invoice Count..."))
                            invoiceCount = Convert.ToInt64(line.Substring(line.IndexOf("Invoice Count...")).Trim());
                        else if (line.Contains("Statement Count."))
                            statementCount = Convert.ToInt64(line.Substring(line.IndexOf("Statement Count.")).Trim());
                        else if (line.Contains("Total Count....."))
                            totalCount = Convert.ToInt64(line.Substring(line.IndexOf("Total Count.....")).Trim());
                    }

                }
            }

            WriteToJobLog(JobLogMessageType.INFO, $"{pageLineNumber} records read");

            if (pageLineNumber > 0)
                WriteHeaderBody();

            //log details
            WriteToJobLog(JobLogMessageType.INFO, $"Company = {company}");
            WriteToJobLog(JobLogMessageType.INFO, $"Bill source = {billSource}");
            WriteToJobLog(JobLogMessageType.INFO, $"Bill date = {billDate}");
            WriteToJobLog(JobLogMessageType.INFO, $"Sort order = {sortOrder}");
            WriteToJobLog(JobLogMessageType.INFO, $"Invoice count = {invoiceCount}");
            WriteToJobLog(JobLogMessageType.INFO, $"Statement count = {statementCount}");
            WriteToJobLog(JobLogMessageType.INFO, $"Total count = {totalCount}");
            WriteToJobLog(JobLogMessageType.INFO, $"Print date = {printDate}");

            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.PROC_UPDATE_LOADS",
                                               new SqlParameter("@pintLoadsID", loadsId),
                                               new SqlParameter("@pvchrCompany", company),
                                               new SqlParameter("@pvchrBillSource", billSource),
                                               new SqlParameter("@psdatBillDate", billDate),
                                               new SqlParameter("@pvchrSortOrder", sortOrder),
                                               new SqlParameter("@pintInvoiceCount", invoiceCount),
                                               new SqlParameter("@pintStatementCount", statementCount),
                                               new SqlParameter("@pintTotalCount", totalCount),
                                               new SqlParameter("@psdatPrintDate", printDate));

            WriteToJobLog(JobLogMessageType.INFO, "Load information updated.");

            //todo: delete file from working directory
            if (pageLineNumber > 0)
            {
                File.Copy(workingFilePath, GetConfigurationKeyValue("OutputDirectory") + workingFileInfo.Name);
                WriteToJobLog(JobLogMessageType.INFO, $"{GetConfigurationKeyValue("OutputDirectory") + workingFileInfo.Name} created.");
            }

            File.Delete(workingFilePath);

            InsertLoadTypes();

            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_PBSReturns_Unscanned_From_Header",
                                            new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, "Innserted into PBSReturn_Unscanned.");
        }

        private void CreateHeaderAndBodyRecords(Int64 loadsId, Int32 headerPageNumber, Int32 groupPageNumber, string printType, string line2Identifier, string invoiceNumber,
                                        string invoiceDate, string billingTerms, string balanceDue, string carrier, string route, string district, string depot,
                                        string truck, string sequence, string nameAddress1, string nameAddress2, string nameAddress3, string nameAddress4,
                                        string barcode, string barcodeReadable, string retailDrawFlag, string printReturnsSheetFlag, string printInvoiceChargeFlag,
                                        string printInvoiceCharge, string printInvoiceCreditFlag, string printInvoiceCredit, string corpSpreadsheetAmount,
                                        string corpSpreadsheetRetailDailyDraw, string corpSpreadsheetRetailDailyDrawCharges, string corpSpreadsheetRetailSundayDrawCharges,
                                        string corpSpreadsheetRetailSundayDraw, string corpSpreadsheetSundayDrawCharges,
                                        string corpSpreadsheetDailyReturns, string corpSpreadsheetDailyReturnCredits, string corpSpreadsheetSundayReturns,
                                        string corpSpreadsheetSundayReturnCredits, string corpSpreadsheetDiscountCredits, string corpSpreadsheetDailyDrawAdjDraw,
                                        string corpSpreadsheetDailyDrawAdjCharges, string corpSpreadsheetSundayDrawAdjDraw, string corpSpreadsheetSundayDrawAdjCharges)
        {
            //save values and reset flags
            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_Header",
                   new SqlParameter("@loads_id", loadsId),
                    new SqlParameter("@header_page_number", headerPageNumber),
                    new SqlParameter("@group_page_number", groupPageNumber),
                    new SqlParameter("@print_type", printType),
                    new SqlParameter("@line_2_identifier", line2Identifier),
                    new SqlParameter("@invoice_number", invoiceNumber),
                    new SqlParameter("@invoice_date", invoiceDate),
                    new SqlParameter("@billing_terms", billingTerms),
                    new SqlParameter("@balance_due", balanceDue),
                    new SqlParameter("@carrier", carrier),
                    new SqlParameter("@route", route),
                    new SqlParameter("@district", district),
                    new SqlParameter("@depot", depot),
                    new SqlParameter("@truck", truck),
                    new SqlParameter("@sequence", sequence),
                    new SqlParameter("@name_address_1", nameAddress1),
                    new SqlParameter("@name_address_2", nameAddress2),
                    new SqlParameter("@name_address_3", nameAddress3),
                    new SqlParameter("@name_address_4", nameAddress4),
                    new SqlParameter("@barcode", barcode),
                    new SqlParameter("@barcode_readable", barcodeReadable),
                    new SqlParameter("@retail_draw_flag", retailDrawFlag),
                    new SqlParameter("@print_returns_sheet_flag", printReturnsSheetFlag),
                    new SqlParameter("@print_invoice_charge_flag", printInvoiceChargeFlag),
                    new SqlParameter("@print_invoice_charge", printInvoiceCharge),
                    new SqlParameter("@print_invoice_credit_flag", printInvoiceCreditFlag),
                    new SqlParameter("@print_invoice_credit", printInvoiceCredit),
                    new SqlParameter("@corporate_spreadsheet_amount", corpSpreadsheetAmount),
                    new SqlParameter("@corporate_spreadsheet_retail_daily_draw", corpSpreadsheetRetailDailyDraw),
                    new SqlParameter("@corporate_spreadsheet_retail_daily_draw_charges", corpSpreadsheetRetailDailyDrawCharges),
                    new SqlParameter("@corporate_spreadsheet_retail_sunday_draw", corpSpreadsheetRetailSundayDraw),
                    new SqlParameter("@corporate_spreadsheet_retail_sunday_draw_charges", corpSpreadsheetRetailSundayDrawCharges),
                    new SqlParameter("@corporate_spreadsheet_daily_returns", corpSpreadsheetDailyReturns),
                    new SqlParameter("@corporate_spreadsheet_daily_return_credits", corpSpreadsheetDailyReturnCredits),
                    new SqlParameter("@corporate_spreadsheet_sunday_returns", corpSpreadsheetSundayReturns),
                    new SqlParameter("@corporate_spreadsheet_sunday_return_credits", corpSpreadsheetSundayReturnCredits),
                    new SqlParameter("@corporate_spreadsheet_discount_credits", corpSpreadsheetDiscountCredits),
                    new SqlParameter("@corporate_spreadsheet_daily_draw_adj_draw", corpSpreadsheetDailyDrawAdjDraw),
                    new SqlParameter("@corporate_spreadsheet_daily_draw_adj_charges", corpSpreadsheetDailyDrawAdjCharges),
                    new SqlParameter("@corporate_spreadsheet_sunday_draw_adj_draw", corpSpreadsheetSundayDrawAdjDraw),
                    new SqlParameter("@corporate_spreadsheet_sunday_draw_adj_charges", corpSpreadsheetSundayDrawAdjCharges));

            if (bodyLineNumber == 0)
            {

            } else
            {

            }

        }

        public override void SetupJob()
        {
            JobName = "PBS Invoices";
            JobDescription = @"";
            AppConfigSectionName = "PBSInvoices";
        }
    }
}
