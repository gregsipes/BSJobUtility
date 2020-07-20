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
                                                                                       new SqlParameter("@pvchrCarrierIdentifier", DBNull.Value));
                //get carrier exceptions
                List<Dictionary<string, object>> carrierExceptions = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Select_Carrier_Exceptions",
                                                                                new SqlParameter("@pvchrCarrier", DBNull.Value));
                //get total identifiers
                List<Dictionary<string, object>> totalIdentifiers = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.PROC_SELECT_TOTAL_IDENTIFIERS",
                                                                                new SqlParameter("@pvchrTotalIdentifier", DBNull.Value));

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

                                CopyAndProcessFile(fileInfo, printTypes);

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
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName, true);
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
                                                                            new SqlParameter("@pvchrAmountDueLabel", DBNull.Value),
                                                                            new SqlParameter("@pflgActiveOnly", 1)).ToList();

            List<Dictionary<string, object>> carrierIdentifiers = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "PROC_SELECT_CARRIER_IDENTIFIERS",
                                                                            new SqlParameter("@pvchrCarrierIdentifier", DBNull.Value)).ToList();

            List<Dictionary<string, object>> carrierExceptions = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "PROC_SELECT_CARRIER_EXCEPTIONS",
                                                                new SqlParameter("@pvchrCarrier", DBNull.Value)).ToList();


            string workingFilePath = GetConfigurationKeyValue("WorkDirectory1") + "carrinv_" + DateTime.Now.ToString("yyMMddhhmmss") + "_" + fileInfo.Name;
            //create a working copy of the file
            File.Copy(fileInfo.FullName, workingFilePath, true);

            FileInfo workingFileInfo = new FileInfo(workingFilePath);


            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(workingFileInfo.FullName).ToList();

            string printDate = "";
            DateTime? invoiceDate = null;
            string invoiceNumber = "";
            decimal? balanceDue = null;
            string route = "";
            string district = "";
            string truck = "";
            string depot = "";
            Int32? sequence = null;
            string billingTerms = "";
            string billSource = "";
            string nameAddress1 = "";
            string nameAddress2 = "";
            string nameAddress3 = "";
            string nameAddress4 = "";
            DateTime? billDate = null;
            string sortOrder = "";
            bool retailDraw = false;
            bool printReturnSheet = false;
            bool printInvoiceCredit = false;
            bool printInvoiceCharge = false;
            Int32 groupPageNumber = 0;

            string carrier = "";
            string pageNumber = "";

            string printTypeIdentifier = "";
            string printTypeName = "";
            string company = "";

            Int32 pageLineNumber = 0;
            Int32 lastPageLineNumber = 0;
            Int32 headerLineNumber = 0;
            Int32 bodyLineNumber = 0;
            Int64 invoiceCount = 0;
            Int64 statementCount = 0;
            Int64 totalCount = 0;

            List<string> headerLines = new List<string>();
            List<string> bodyLines = new List<string>();

            decimal printInvoiceChargeTotal = 0;
            decimal printInvoiceCreditTotal = 0;

            bool checkIdentifiers = true;
            bool checkRouteSuffix = false;


            foreach (string line in fileContents)
            {

                //if (line != null && (line.Trim().Length > 0 | line.Contains("\f")))
                //{

                if (line.StartsWith("PRINT DATE:"))
                {
                    printDate = line.Replace("PRINT DATE:", "").Trim(); //this is the last statement that we will process in the file
                    break;
                }
                else if (line.Contains("\f"))
                {
                    headerLineNumber += 100;

                    //testing
                    //if (headerLineNumber > 53200)
                    //{
                    //    var x = 2 + 1;
                    //}

                    CreateHeaderAndBodyRecords(loadsId, headerLineNumber, groupPageNumber, printTypeName, printTypeIdentifier,
                                                invoiceNumber, invoiceDate, billingTerms, balanceDue.ToString() ?? "", carrier, route,
                                                district, depot, truck, sequence.ToString() ?? "", nameAddress1, nameAddress2, nameAddress3, nameAddress4,
                                                retailDraw, printReturnSheet, printInvoiceCharge, printInvoiceChargeTotal, printInvoiceCredit, printInvoiceCreditTotal, bodyLines);


                    pageLineNumber = 0;

                    printTypeIdentifier = "";
                    printTypeName = "";
                    printDate = "";
                    invoiceDate = null;
                    invoiceNumber = "";
                    balanceDue = null;
                    route = "";
                    district = "";
                    truck = "";
                    depot = "";
                    sequence = null;
                    billingTerms = "";
                    billSource = "";
                    nameAddress1 = "";
                    nameAddress2 = "";
                    nameAddress3 = "";
                    nameAddress4 = "";
                    billDate = null;
                    sortOrder = "";
                //    retailDraw = false;
               //     printReturnSheet = false;
                //    printInvoiceCredit = false;
               //    printInvoiceCharge = false;

                    checkIdentifiers = true;
                    checkRouteSuffix = false;
                }


                if ((pageLineNumber == 0 && line.Trim() != "") | pageLineNumber > 0)
                {
                    pageLineNumber++;

                    switch (pageLineNumber)
                    {
                        case 1:
                            if (line.Contains("** Value"))
                                pageLineNumber--; //does this condition ever get hit?
                            else
                            {
                                lastPageLineNumber++;
                                pageNumber = line.Replace("PAGE:", "").Trim();

                                Int32 page;
                                if (Int32.TryParse(pageNumber, out page))
                                    groupPageNumber = page;
                                else
                                    groupPageNumber = lastPageLineNumber;

                            }

                            if (groupPageNumber == 1)
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
                                if (printType["line_2_identifier_1"].ToString() != "" && line.ToLower().Contains(printType["line_2_identifier_1"].ToString().ToLower()))
                                {
                                    printTypeIdentifier = printType["line_2_identifier_1"].ToString();
                                    printTypeName = printType["print_type"].ToString();

                                    invoiceDate = DateTime.Parse(line.Substring(line.IndexOf(printTypeIdentifier.ToUpper())).Replace(printTypeIdentifier.ToUpper(), "").Replace("DATE:", "").Trim());
                                    checkIdentifiers = !bool.Parse(printType["do_not_check_carrier_identifiers_flag"].ToString());
                                    checkRouteSuffix = bool.Parse(printType["check_route_suffix_flag"].ToString());

                                    break;
                                }

                                if (printType["line_2_identifier_2"].ToString() != "" && line.ToLower().Contains(printType["line_2_identifier_2"].ToString().ToLower()))
                                {
                                    printTypeIdentifier = printType["line_2_identifier_2"].ToString();
                                    printTypeName = printType["print_type"].ToString();

                                    invoiceDate = DateTime.Parse(line.Substring(line.IndexOf(printTypeIdentifier.ToUpper())).Replace(printTypeIdentifier.ToUpper(), "").Replace("DATE:", "").Trim());
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
                            billingTerms = line.Trim(); //.Substring(line.IndexOf("BILLING TERMS:")).Replace("BILLING TERMS:", "").Trim();
                            break;
                        case 8:
                            //check for any possible label values (BALANCE DUE, TOTAL BALANCE DUE, TOTAL OUTSTANDING BALANCE) 
                            foreach (Dictionary<string, object> amountLabel in amountLabels)
                            {
                                if (line.Contains(amountLabel["amount_due_label"].ToString()))
                                {
                                    string x = line.Replace(amountLabel["amount_due_label"].ToString(), "").Replace(":", "").Replace(",", "").Trim();
                                    balanceDue = decimal.Parse(line.Replace(amountLabel["amount_due_label"].ToString(), "").Replace(":", "").Replace(",", "").Trim());
                                    break;
                                }
                            }
                            break;
                        case 10:
                            nameAddress1 = line.Trim().Substring(0, 40).Trim();
                            carrier = line.Trim().Replace(nameAddress1, "").Replace("ACCOUNT     :", "").Trim();

                            //Check if carrier/print type combinations is in carrier exceptions array.
                            if (carrierExceptions != null && carrierExceptions.Count() > 0 && printTypeIdentifier != "")
                            {
                                foreach (Dictionary<string, object> exception in carrierExceptions)
                                {
                                    if (exception["carrier"].ToString() == carrier)
                                    {
                                        if (exception["print_type_1"].ToString() != "" && printTypeIdentifier == exception["line_2_identifier_1"].ToString())  //print_type_1 is never empty, maybe this is something can happen in the UI?
                                        {
                                            checkIdentifiers = false;
                                            checkRouteSuffix = false;
                                            printTypeName = exception["print_type_1"].ToString();
                                        }
                                    }
                                    else if (exception["print_type_2"].ToString() != "" && printTypeIdentifier == exception["line_2_identifier_2"].ToString())  //print_type_2 is never empty, maybe this is something can happen in the UI?
                                    {
                                        checkIdentifiers = false;
                                        checkRouteSuffix = false;
                                        printTypeName = exception["print_type_2"].ToString();
                                    }
                                    else if (exception["print_type_3"].ToString() != "" && printTypeIdentifier == exception["line_2_identifier_3"].ToString())  //this condition is never currently hit
                                    {
                                        checkIdentifiers = false;
                                        checkRouteSuffix = false;
                                        printTypeName = exception["print_type_3"].ToString();
                                    }
                                }
                            }

                            //Check the carrier to determine print type.
                            if (checkIdentifiers && carrierIdentifiers != null)
                            {
                                foreach (Dictionary<string, object> carrierIdentifier in carrierIdentifiers)
                                {
                                    if (carrier.StartsWith(carrierIdentifier["carrier_identifier"].ToString()))
                                    {
                                        printTypeName = carrierIdentifier["print_type"].ToString();
                                        checkRouteSuffix = false;
                                        break;
                                    }
                                }
                            }

                            break;
                        case 11:
                            if (printTypeName == printTypes.Where(p => Boolean.Parse(p["total_flag"].ToString()) == true).Select(p => p["print_type"].ToString()).FirstOrDefault())
                                company = line.Substring(0, 40).Replace("COMPANY    :", "").Trim(); 
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
                                    if (bool.Parse(printType["check_route_suffix_flag"].ToString()) != false &
                                        ((bool.Parse(printType["route_suffix_alpha_flag"].ToString()) != false & route.Substring(route.Length - 2, 1).All(char.IsNumber))
                                                 | (bool.Parse(printType["route_suffix_alpha_flag"].ToString()) == false & !route.Substring(route.Length - 2, 1).All(char.IsNumber))) &
                                         (printType["line_2_identifier_1"].ToString() == printTypeIdentifier | printType["line_2_identifier_2"].ToString() == printTypeIdentifier))
                                    {
                                        printTypeName = printType["print_type"].ToString();
                                        break;
                                    }
                                }
                            }

                            break;
                        case 12:
                            if (printTypeName == printTypes.Where(p => Boolean.Parse(p["total_flag"].ToString()) == true).Select(p => p["print_type"].ToString()).FirstOrDefault())
                                billSource = line.Substring(0, 40).Replace("BILL SOURCE:", "").Trim();
                            else
                                nameAddress3 = line.Trim().Substring(0, 40).Trim();

                            break;
                        case 13:
                            if (printTypeName == printTypes.Where(p => Boolean.Parse(p["total_flag"].ToString()) == true).Select(p => p["print_type"].ToString()).FirstOrDefault())
                            {
                                DateTime date;
                                DateTime.TryParse(line.Substring(0, 40).Replace("BILL DATE  :", "").Trim(), out date);
                                billDate = date;
                            }
                            else
                                nameAddress4 = line.Substring(0, 40).Trim();

                            break;
                        case 14:
                            if (printTypeName == printTypes.Where(p => Boolean.Parse(p["total_flag"].ToString()) == true).Select(p => p["print_type"].ToString()).FirstOrDefault())
                                sortOrder = line.Substring(0, 40).Replace("SORT ORDER :", "").Trim();

                            break;
                    }

                    if (printTypeName == printTypes.Where(p => Boolean.Parse(p["total_flag"].ToString()) == true).Select(p => p["print_type"].ToString()).FirstOrDefault())
                    {
                        if (pageLineNumber >= 12 & pageLineNumber <= 15)
                        {
                            if (line.Contains("DISTRICT    :"))
                                district = line.Substring(line.IndexOf("DISTRICT    :")).Replace("DISTRICT    :", "").Trim();
                            else if (line.Contains("TRUCK       :"))
                                truck = line.Substring(line.IndexOf("TRUCK       :")).Replace("TRUCK       :", "").Trim();
                            else if (line.Contains("DEPOT       :"))
                                depot = line.Substring(line.IndexOf("DEPOT       :")).Replace("DEPOT       :", "").Trim();
                            else if (line.Contains("SEQUENCE    :"))
                                sequence = Convert.ToInt32(line.Substring(line.IndexOf("SEQUENCE    :")).Replace("SEQUENCE    :", "").Trim());
                        }
                    }

                    if (pageLineNumber > 0 && pageLineNumber <= 16)
                    {
                        //headerLineNumber++;
                        headerLines.Add(line);
                    }
                    else if (pageLineNumber > 16 && line.Length > 18)
                    {
                        bodyLineNumber++;
                        bodyLines.Add(line);

                        if (line.Contains("RETAIL DAILY DRAW") | line.Contains("RETAIL SUNDAY DRAW") | line.Contains("CORP STORE DELIVERY CREDIT") |
                            line.Contains("DIRECT BILL DELIVERY CREDIT") | line.Contains("RETURN CREDITS") | line.Contains("USA RETAIL HONOR BOX CHARGE"))
                        {
                            printReturnSheet = true;
                        }

                        if (line.Substring(17).Contains("RETAIL DAILY DRAW") | line.Substring(17).Contains("RETAIL SUNDAY DRAW"))
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

                    if (printTypeName == printTypes.Where(p => Boolean.Parse(p["total_flag"].ToString()) == true).Select(p => p["print_type"].ToString()).FirstOrDefault())
                    {
                        if (line.Contains("Invoice Count................."))
                            invoiceCount = Convert.ToInt64(line.Replace("Invoice Count.................", "").Trim().Substring(0, 14));
                        else if (line.Contains("Statement Count..............."))
                            statementCount = Convert.ToInt64(line.Replace("Statement Count...............", "").Trim().Substring(0, 14));
                        else if (line.Contains("Total Count..................."))
                            totalCount = Convert.ToInt64(line.Replace("Total Count...................", "").Trim().Substring(0, 14));
                    }

                }
            }

            WriteToJobLog(JobLogMessageType.INFO, $"{pageLineNumber} records read");

            if (pageLineNumber > 0)
            {
                headerLineNumber += 100;

                ////testing
                //if (headerLineNumber == 53400)
                //{
                //    var x = 2 + 1;
                //}

                CreateHeaderAndBodyRecords(loadsId, headerLineNumber, groupPageNumber, printTypeName, printTypeIdentifier,
                            invoiceNumber, invoiceDate, billingTerms, balanceDue.Value.ToString(), carrier, route,
                            district, depot, truck, sequence.Value.ToString(), nameAddress1, nameAddress2, nameAddress3, nameAddress4,
                            retailDraw, printReturnSheet, printInvoiceCharge, printInvoiceChargeTotal, printInvoiceCredit, printInvoiceCreditTotal,
                            bodyLines);
            }

            //log details
            WriteToJobLog(JobLogMessageType.INFO, $"Company = {company}");
            WriteToJobLog(JobLogMessageType.INFO, $"Bill source = {billSource}");
            WriteToJobLog(JobLogMessageType.INFO, $"Bill date = {billDate.ToString()}");
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
                File.Copy(workingFilePath, GetConfigurationKeyValue("OutputDirectory") + workingFileInfo.Name, true);
                WriteToJobLog(JobLogMessageType.INFO, $"{GetConfigurationKeyValue("OutputDirectory") + workingFileInfo.Name} created.");
            }

            File.Delete(workingFilePath);

            InsertLoadsPrintTypes(loadsId, printTypes);

            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_PBSReturns_Unscanned_From_Header",
                                            new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, "Innserted into PBSReturn_Unscanned.");
        }

        private void CreateHeaderAndBodyRecords(Int64 loadsId, Int32 headerPageNumber, Int32 groupPageNumber, string printType, string line2Identifier, string invoiceNumber,
                                        DateTime? invoiceDate, string billingTerms, string balanceDue, string carrier, string route, string district, string depot,
                                        string truck, string sequence, string nameAddress1, string nameAddress2, string nameAddress3, string nameAddress4,
                                        bool retailDrawFlag, bool printReturnsSheetFlag, bool printInvoiceChargeFlag,
                                        decimal printInvoiceCharge, bool printInvoiceCreditFlag, decimal printInvoiceCredit, List<string> bodyLines)
        {


            //save values and reset flags
            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_Header",
                   new SqlParameter("@loads_id", loadsId),
                    new SqlParameter("@header_page_number", headerPageNumber),
                    new SqlParameter("@group_page_number", groupPageNumber),
                    new SqlParameter("@print_type", printType),
                    new SqlParameter("@line_2_identifier", line2Identifier),
                    new SqlParameter("@invoice_number", invoiceNumber),
                    new SqlParameter("@invoice_date", invoiceDate.HasValue ? invoiceDate.Value : (object)DBNull.Value),
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
                    new SqlParameter("@barcode", DBNull.Value),
                    new SqlParameter("@barcode_readable", DBNull.Value),
                    new SqlParameter("@retail_draw_flag", retailDrawFlag),
                    new SqlParameter("@print_returns_sheet_flag", printReturnsSheetFlag),
                    new SqlParameter("@print_invoice_charge_flag", printInvoiceChargeFlag),
                    new SqlParameter("@print_invoice_charge", printInvoiceCharge),
                    new SqlParameter("@print_invoice_credit_flag", printInvoiceCreditFlag),
                    new SqlParameter("@print_invoice_credit", printInvoiceCredit),
                    new SqlParameter("@corporate_spreadsheet_amount", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_retail_daily_draw", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_retail_daily_draw_charges", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_retail_sunday_draw", "0"),
                    new SqlParameter("@corporate_spreadsheet_retail_sunday_draw_charges", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_daily_returns", "0"),
                    new SqlParameter("@corporate_spreadsheet_daily_return_credits", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_sunday_returns", "0"),
                    new SqlParameter("@corporate_spreadsheet_sunday_return_credits", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_discount_credits", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_daily_draw_adj_draw", "0"),
                    new SqlParameter("@corporate_spreadsheet_daily_draw_adj_charges", "$0.00"),
                    new SqlParameter("@corporate_spreadsheet_sunday_draw_adj_draw", "0"),
                    new SqlParameter("@corporate_spreadsheet_sunday_draw_adj_charges", "$0.00"));

            if (bodyLines.Count() == 0)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_Body",
                                           new SqlParameter("@loads_id", loadsId),
                                           new SqlParameter("@header_page_number", headerPageNumber),
                                           new SqlParameter("@body_line_number", 1),
                                           new SqlParameter("@body_text", ""));
            }
            else
            {
                Int32 lineNumber = 0;
                foreach (string line in bodyLines)
                {
                    lineNumber++;

                    ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Insert_Body",
                           new SqlParameter("@loads_id", loadsId),
                           new SqlParameter("@header_page_number", headerPageNumber),
                           new SqlParameter("@body_line_number", lineNumber),
                           new SqlParameter("@body_text", line));
                }
            }

        }

        private void InsertLoadsPrintTypes(Int64 loadsId, List<Dictionary<string, object>> printTypes)
        {
            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.PROC_INSERT_LOADS_PRINT_TYPES",
                                new SqlParameter("@loads_id", loadsId));

            WriteToJobLog(JobLogMessageType.INFO, "Set header count for each print type.");

            List<Dictionary<string, object>> loadsPrintTypes = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Select_Loads_Print_Types",
                                                                            new SqlParameter("@loads_id", loadsId)).ToList();

            foreach (Dictionary<string, object> loadPrintType in loadsPrintTypes)
            {
                foreach (Dictionary<string, object> printType in printTypes)
                {
                    if (printType["print_type"].ToString() == loadPrintType["print_type"].ToString())
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "dbo.Proc_Update_Load_Print_Types",
                                 new SqlParameter("@pintLoadsID", loadsId),
                                new SqlParameter("@print_Type", printType["print_type"].ToString()),
                                new SqlParameter("@line_2_identifier_1", printType["line_2_identifier_1"].ToString()),
                                new SqlParameter("@line_2_identifier_2", printType["line_2_identifier_2"].ToString()),
                                new SqlParameter("@do_not_check_carrier_identifiers_flag", printType["do_not_check_carrier_identifiers_flag"].ToString()),
                                new SqlParameter("@check_route_suffix_flag", printType["check_route_suffix_flag"].ToString()),
                                new SqlParameter("@route_suffix_alpha_flag", printType["route_suffix_alpha_flag"].ToString()));

                        break;
                    }
                }
            }


            WriteToJobLog(JobLogMessageType.INFO, "Load specific options noted for each print type.");
        }

        public override void SetupJob()
        {
            JobName = "PBS Invoices";
            JobDescription = @"Parses multiple fixed width files";
            AppConfigSectionName = "PBSInvoices";
        }
    }
}
