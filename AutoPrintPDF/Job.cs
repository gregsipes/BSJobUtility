using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace AutoPrintPDF
{
    public class Job : JobBase
    {
        public string Version { get; set; }

        public override void SetupJob()
        {
            JobName = "AutoPrintPDF";
            JobDescription = "TODO";
            AppConfigSectionName = "AutoPrintPDF";

        }

        public override void ExecuteJob()
        {
            try
            {
                switch (Version)
                {
                    case "OfficePay":
                    case "AutoRenew":
                        AutoRenewOrOfficePay();
                        break;
                    case "PBSInvoices":
                        PBSInvoices();
                        break;
                    case "PBSInvoicesByCarrierID":
                        //todo:
                        break;
                    default:
                        throw new Exception("Unknown version");
                }
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void AutoRenewOrOfficePay()
        {
            string description = "renewal";

            if (Version == "AutoRenew")
                description = "autorenew";

            WriteToJobLog(JobLogMessageType.INFO, $"Determining {description} notices to send to .pdf");

            List<Dictionary<string, object>> loads = new List<Dictionary<string, object>>();

            if (Version == "AutoRenew")
                loads = ExecuteSQL(DatabaseConnectionStringNames.AutoRenew, "Proc_Select_Loads_For_AutoPrint_To_PDF").ToList();
            else
                loads = ExecuteSQL(DatabaseConnectionStringNames.OfficePay, "Proc_Select_Loads_For_AutoPrint_To_PDF").ToList();

            if (loads == null || loads.Count() == 0)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"No {description} notices to create .pdf's for exist in database");
                return;
            }

            //todo: install tru type font?


            WriteToJobLog(JobLogMessageType.INFO, "Creating .pdf's in work directory as {subscription_number}{MMddyyyy}INVOICE.pdf");

            foreach (Dictionary<string, object> load in loads)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Retrieving {description} notices for loads_id {load["loads_id"].ToString()}");

                List<Dictionary<string, object>> results = new List<Dictionary<string, object>>();

                if (Version == "AutoRenew")
                    results = ExecuteSQL(DatabaseConnectionStringNames.AutoRenew, "Proc_Select_For_AutoRenew",
                                                        new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                                        new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()),
                                                        new SqlParameter("@pflgOnlyWithEmailAddress", false),
                                                        new SqlParameter("@pflgReport", true),
                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword"))).ToList();
                else
                    results = ExecuteSQL(DatabaseConnectionStringNames.OfficePay, "Proc_Select_For_Office_Pay_Bills",
                                                        new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                                        new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()),
                                                        new SqlParameter("@pvchrRenewalType", null),
                                                        new SqlParameter("@pvchrRenewalNumber", 0),
                                                        new SqlParameter("@pflgZero", 0),
                                                        new SqlParameter("@pflgDuplicate", 0),
                                                        new SqlParameter("@pflgAutoPrintToPDF", 1),
                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword"))).ToList();


                if (results == null || results.Count() == 0)
                {
                    WriteToJobLog(JobLogMessageType.INFO, $"No {description} notices exist for this loads_id");
                    return;
                }

                Int32 subDirectoryCount = 1;

                //create output path. ex - \\omaha\AutoPrintPDF_AutoRenew\20201021090010_3FBFFF3498914385BE6B2E0E3919E046\1\
                string baseOutputDirectory = GetConfigurationKeyValue("WorkDirectory") + Version + "\\" + DateTime.Now.ToString("yyyyMMddhhmmss") + "_" + Guid.NewGuid().ToString().Replace("-", "") + "\\";

                if (!Directory.Exists(baseOutputDirectory))
                    Directory.CreateDirectory(baseOutputDirectory);

                WriteToJobLog(JobLogMessageType.INFO, $"{results.Count()} {description} notices to be created for renewal run date(s) {load["renewal_run_dates"].ToString()}");
                WriteToJobLog(JobLogMessageType.INFO, $".pdf's being created in {baseOutputDirectory}");

                Int32 totalCounter = 0;

                foreach (Dictionary<string, object> result in results)
                {
                    totalCounter++;

                   string outputDirectory = baseOutputDirectory + subDirectoryCount.ToString() + "\\";

                    if (!Directory.Exists(outputDirectory))
                        Directory.CreateDirectory(outputDirectory);

                    string outputFileName = result["subscription_number_without_check_digit"].ToString() + Convert.ToDateTime(result["renewal_run_date"].ToString()).ToString("MMddyyyy") + "INVOICE.pdf";

                    if (Version == "AutoRenew")
                    {
                        //todo: call reports here



                        //create record in AutoPrintPDF database
                        ExecuteNonQuery(DatabaseConnectionStringNames.AutoPrintPDF, "Proc_Insert_AutoRenew",
                                                    new SqlParameter("@pvchrFileName", outputFileName),
                                                    new SqlParameter("@psdatRenewalDate", Convert.ToDateTime(result["renewal_run_date"].ToString()).ToShortDateString()));
                    }

                    else
                    {
                        //todo: call reports here


                        //create record in AutoPrintPDF database
                        ExecuteNonQuery(DatabaseConnectionStringNames.AutoPrintPDF, "Proc_Insert_OfficePay",
                                                    new SqlParameter("@pvchrFileName", outputFileName),
                                                    new SqlParameter("@psdatRenewalDate", Convert.ToDateTime(result["renewal_run_date"].ToString()).ToShortDateString()));
                    }


                    //log every 60th file
                    if (totalCounter % 60 == 0)
                    {
                        WriteToJobLog(JobLogMessageType.INFO, $"{totalCounter} {description} notices created in work directory...");

                    }

                    //after 9,900 files, create a new sub directory
                    if (totalCounter % 9900 == 0)
                         subDirectoryCount++;


                    //copy files to cmpdf directory
                    File.Copy(outputDirectory + outputFileName, GetConfigurationKeyValue("CopyDirectory") + outputFileName);

                    WriteToJobLog(JobLogMessageType.INFO, $"File copied to {GetConfigurationKeyValue("CopyDirectory") + outputFileName}");

                }

                //run update sproc
                if (Version == "AutoRenew")
                    ExecuteNonQuery(DatabaseConnectionStringNames.AutoRenew, "Proc_Insert_Loads_Successful_AutoPrint_To_PDF",
                                            new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                            new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()));
                else
                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "Proc_Insert_Loads_Successful_AutoPrint_To_PDF",
                                            new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                            new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()));


                //remove work directory files
                WriteToJobLog(JobLogMessageType.INFO, $"Deleting work directory {baseOutputDirectory}");
                Directory.Delete(baseOutputDirectory, true);
            }

            WriteToJobLog(JobLogMessageType.INFO, "AutoRenewOrOfficePay processing completed");

        }

        private void PBSInvoices()
        {
            WriteToJobLog(JobLogMessageType.INFO, "Determining latest invoice date");

            List<Dictionary<string, object>> loads = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Loads_Bill_Dates_Pages").ToList();

            if (loads == null || loads.Count() == 0)
            {
                WriteToJobLog(JobLogMessageType.INFO, "No invoice dates for which .pdf is to be created exist in database");
                return;
            }

            foreach(Dictionary<string, object> load in loads)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Found loads_id {load["loads_id"].ToString()}");
                WriteToJobLog(JobLogMessageType.INFO, $"Retrieving invoices for {load["bill_date"].ToString()}");

                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Header_Body_By_Bill_Date_No_Additional_Copies",
                                                                        new SqlParameter("@psdatBillDate", load["bill_date"].ToString())).ToList();

                if (results.Count() == 0)
                    WriteToJobLog(JobLogMessageType.INFO, $"No invoices exist for {load["bill_date"].ToString()}");
                else
                {
                    string outputFile = GetConfigurationKeyValue("PBSInvoiceDirectory") + Convert.ToDateTime(load["bill_date"].ToString()).ToShortDateString() + ".pdf";

                    //if the file already exists, delete it
                    if (File.Exists(outputFile))
                        File.Delete(outputFile);

                    WriteToJobLog(JobLogMessageType.INFO, $"Sending invoices to {outputFile}");

                    //todo: call reports here

                    DeleteTemp();

                    //run update sproc
                    ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "Proc_Update_Loads_Successful_AutoPrint_to_PDF_Flag",
                                                    new SqlParameter("@pintLoadsID", load["loads_id"].ToString()));

                    WriteToJobLog(JobLogMessageType.INFO, $"{outputFile} successfully created");

                }

            }
        }

        private void PBSInvoicesByCarrierID()
        {
            WriteToJobLog(JobLogMessageType.INFO, "Determining latest invoice date");

            List<Dictionary<string, object>> loads = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Loads_Bill_Dates_Pages_For_AutoPrint_to_PDF_By_CarrierID").ToList();

            if (loads == null || loads.Count() == 0)
            {
                WriteToJobLog(JobLogMessageType.INFO, "No invoice dates for which .pdf is to be created exist in database");
                return;
            }

            //todo: install font?

            foreach(Dictionary<string, object> load in loads)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Found loads_id {load["loads_id"].ToString()}");
                WriteToJobLog(JobLogMessageType.INFO, $"Bill Date = {load["bill_date"].ToString()}");
                WriteToJobLog(JobLogMessageType.INFO, $"Bill Source = {load["bill_source"].ToString()}");

                List<Dictionary<string, object>> carriers = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Header_Distinct_Carrier",
                                                                                        new SqlParameter("@psdatBillDate", load["bill_date"].ToString()),
                                                                                        new SqlParameter("@pvchrBillSource", load["bill_source"].ToString())).ToList();

                if (carriers == null || carriers.Count() == 0)
                    WriteToJobLog(JobLogMessageType.INFO, "No invoices exist for this load, bill date, bill source & print type");
                else
                {
                    foreach (Dictionary<string, object> carrier in carriers)
                    {

                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Header_Body_By_Bill_Date_No_Additional_Copies_By_CarrierID",
                                                                                        new SqlParameter("@psdatBillDate", load["bill_date"].ToString()),
                                                                                        new SqlParameter("@pvchrBillSource", load["bill_source"].ToString()),
                                                                                        new SqlParameter("@pvchrCarrier", carrier["carrier"].ToString())).ToList();

                        //string outputFileName = 

                        //todo:
                    }
                }
            }

        }

        private void DeleteTemp()
        {
            //GDS - I'm not sure why this is needed here or why these temp files are getting created. From
            //the current logs, it looks like this code is still in use, so it has been carried over

            List<string> tempFiles = Directory.GetFiles(Path.GetTempPath(), "ctm*.tmp").ToList();

            foreach(string tempFile in tempFiles)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Delete temp file {tempFile}");
                File.Delete(tempFile);
            }

        }

    }
}
