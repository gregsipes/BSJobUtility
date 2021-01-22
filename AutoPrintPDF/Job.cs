using BSJobBase;
using CrystalDecisions.Shared;
using Reporting;
using System;
using System.Collections.Generic;
using System.Data;
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

        public override void SetupJob()
        {
            JobName = "AutoPrintPDF";
            JobDescription = "Generates PDF files for renewal notices and carrier invoices";
            AppConfigSectionName = "AutoPrintPDF";

        }

        public override void ExecuteJob()
        {
            try
            {
                AutoRenewOrOfficePay("AutoRenew");
                AutoRenewOrOfficePay("OfficePay");
                PBSInvoices();
                PBSInvoicesByCarrierID();
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void AutoRenewOrOfficePay(string version)
        {
           // WriteToJobLog(JobLogMessageType.INFO, $"Determining {version} notices to send to .pdf");

            List<Dictionary<string, object>> loads = new List<Dictionary<string, object>>();

            if (version == "AutoRenew")
                loads = ExecuteSQL(DatabaseConnectionStringNames.AutoRenew, "Proc_Select_Loads_For_AutoPrint_To_PDF").ToList();
            else
                loads = ExecuteSQL(DatabaseConnectionStringNames.OfficePay, "Proc_Select_Loads_For_AutoPrint_To_PDF").ToList();

            if (loads == null || loads.Count() == 0)
                return;


            foreach (Dictionary<string, object> load in loads)
            {
                //WriteToJobLog(JobLogMessageType.INFO, $"Retrieving {version} notices for loads_id {load["loads_id"].ToString()}");

                 SqlDataReader reader = null;

                if (version == "AutoRenew")
                    reader = ExecuteSQLReturnDataReader(DatabaseConnectionStringNames.AutoRenew, CommandType.StoredProcedure, "Proc_Select_For_AutoRenew",
                                                        new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                                        new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()),
                                                        new SqlParameter("@pflgOnlyWithEmailAddress", false),
                                                        new SqlParameter("@pflgReport", true),
                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")));
                else
                    reader = ExecuteSQLReturnDataReader(DatabaseConnectionStringNames.OfficePay, CommandType.StoredProcedure, "Proc_Select_For_Office_Pay_Bills",
                                                        new SqlParameter("@pintLoadsID", load["loads_id"].ToString()),
                                                        new SqlParameter("@pvchrPublicationName", load["publication_name"].ToString()),
                                                        new SqlParameter("@pvchrRenewalType", DBNull.Value),
                                                        new SqlParameter("@pvchrRenewalNumber", "0"),
                                                        new SqlParameter("@pflgZero", false),
                                                        new SqlParameter("@pflgDuplicate", false),
                                                        new SqlParameter("@pflgAutoPrintToPDF", true),
                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")));


                if (reader == null || !reader.HasRows)
                    return;

                WriteToJobLog(JobLogMessageType.INFO, "Creating .pdf's in work directory as {subscription_number}{MMddyyyy}INVOICE.pdf");

                Int32 subDirectoryCount = 1;

                //create output path. ex - \\omaha\AutoPrintPDF_AutoRenew\20201021090010_3FBFFF3498914385BE6B2E0E3919E046\1\
                string baseOutputDirectory = GetConfigurationKeyValue("WorkDirectory") + version + "\\" + DateTime.Now.ToString("yyyyMMddhhmmss") + "_" + Guid.NewGuid().ToString().Replace("-", "") + "\\";

                if (!Directory.Exists(baseOutputDirectory))
                    Directory.CreateDirectory(baseOutputDirectory);

                WriteToJobLog(JobLogMessageType.INFO, $".pdf's being created in {baseOutputDirectory}");

                Int32 totalCounter = 0;

                while (reader.Read())
                {
                    Dictionary<string, object> result = ConvertDataReaderToDictionary(reader);

                    totalCounter++;

                    string outputDirectory = baseOutputDirectory + subDirectoryCount.ToString() + "\\";

                    if (!Directory.Exists(outputDirectory))
                        Directory.CreateDirectory(outputDirectory);

                    string outputFileName = reader["subscription_number_without_check_digit"].ToString() + Convert.ToDateTime(reader["renewal_run_date"].ToString()).ToString("MMddyyyy") + "INVOICE.pdf";

                    if (version == "AutoRenew")
                    {
                        //generate and save reports
                        if (reader["report_name"].ToString() == "rptAutoRenew")
                        {
                            using (rptAutoRenew report = new rptAutoRenew())
                            {
                                report.SetDataSource((IDataReader)reader);
                                report.ExportToDisk(ExportFormatType.PortableDocFormat, outputDirectory + outputFileName);
                            }

                        }
                        else if (reader["report_name"].ToString() == "rptAutoRenewPrintDigital")
                        {
                            using (rptAutoRenewPrintDigital report = new rptAutoRenewPrintDigital())
                            {
                                report.SetDataSource((IDataReader)reader);
                                report.ExportToDisk(ExportFormatType.PortableDocFormat, outputDirectory + outputFileName);
                            }

                        }
                        else if (reader["report_name"].ToString() == "rptAutoRenewSun")
                        {
                            using (rptAutoRenewSun report = new rptAutoRenewSun())
                            {
                                report.SetDataSource((IDataReader)reader);
                                report.ExportToDisk(ExportFormatType.PortableDocFormat, outputDirectory + outputFileName);
                            }
                        }

                        //create record in AutoPrintPDF database
                        ExecuteNonQuery(DatabaseConnectionStringNames.AutoPrintPDF, "Proc_Insert_AutoRenew",
                                                    new SqlParameter("@pvchrFileName", outputDirectory + outputFileName),
                                                    new SqlParameter("@psdatRenewalDate", Convert.ToDateTime(result["renewal_run_date"].ToString()).ToShortDateString()));
                    }

                    else
                    {
                        //generate and save reports
                        if (reader["report_name"].ToString() == "rptOfficePayPrintDigital")
                        {
                            using (rptOfficePayPrintDigital report = new rptOfficePayPrintDigital())
                            {
                                report.SetDataSource((IDataReader)reader);
                                report.ExportToDisk(ExportFormatType.PortableDocFormat, outputDirectory + outputFileName);
                            }
                        }
                        else if (reader["report_name"].ToString() == "rptOfficePaySun")
                        {
                            using (rptOfficePaySun report = new rptOfficePaySun())
                            {
                                report.SetDataSource((IDataReader)reader);
                                report.ExportToDisk(ExportFormatType.PortableDocFormat, outputDirectory + outputFileName);
                            }
                        }

                        //create record in AutoPrintPDF database
                        ExecuteNonQuery(DatabaseConnectionStringNames.AutoPrintPDF, "Proc_Insert_OfficePay",
                                                    new SqlParameter("@pvchrFileName", outputDirectory + outputFileName),
                                                    new SqlParameter("@psdatRenewalDate", Convert.ToDateTime(result["renewal_run_date"].ToString()).ToShortDateString()));
                    }


                    //log every 60th file
                    if (totalCounter % 60 == 0)
                    {
                        WriteToJobLog(JobLogMessageType.INFO, $"{totalCounter} {version} notices created in work directory...");

                    }

                    //after 9,900 files, create a new sub directory
                    //GDS - this will never be the case. Crystal Reports will error out long before 9,900 are generated
                    if (totalCounter % 9900 == 0)
                        subDirectoryCount++;


                    //copy files to cmpdf directory 
                    File.Copy(outputDirectory + outputFileName, GetConfigurationKeyValue("CopyDirectory") + outputFileName);

                    WriteToJobLog(JobLogMessageType.INFO, $"File copied to {GetConfigurationKeyValue("CopyDirectory") + outputFileName}");

                }

                //run update sproc
                if (version == "AutoRenew")
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
           // WriteToJobLog(JobLogMessageType.INFO, "Determining latest invoice date");

            List<Dictionary<string, object>> loads = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Loads_Bill_Dates_Pages").ToList();

            if (loads == null || loads.Count() == 0)
                return;

            foreach (Dictionary<string, object> load in loads)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Found loads_id {load["loads_id"].ToString()}");
                WriteToJobLog(JobLogMessageType.INFO, $"Retrieving invoices for {load["bill_date"].ToString()}");

                SqlDataReader results = ExecuteSQLReturnDataReader(DatabaseConnectionStringNames.PBSInvoices, CommandType.StoredProcedure, "Proc_Select_Header_Body_By_Bill_Date_No_Additional_Copies",
                                                                       new SqlParameter("@psdatBillDate", load["bill_date"].ToString()));  



                if (!results.HasRows)
                    WriteToJobLog(JobLogMessageType.INFO, $"No invoices exist for {load["bill_date"].ToString()}");
                else
                {
                    string outputFile = GetConfigurationKeyValue("PBSInvoiceDirectory") + Convert.ToDateTime(load["bill_date"].ToString()).ToString("yyyy-MM-dd") + ".pdf";

                    //if the file already exists, delete it
                    if (File.Exists(outputFile))
                        File.Delete(outputFile);

                    WriteToJobLog(JobLogMessageType.INFO, $"Sending invoices to {outputFile}");

                    using (rptInvoices report = new rptInvoices())
                    {
                        report.SetDataSource((IDataReader)results);
                        report.ExportToDisk(ExportFormatType.PortableDocFormat, outputFile);
                    }

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
           // WriteToJobLog(JobLogMessageType.INFO, "Determining latest invoice date");

            List<Dictionary<string, object>> loads = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoices, "Proc_Select_Loads_Bill_Dates_Pages_For_AutoPrint_to_PDF_By_CarrierID").ToList();

            if (loads == null || loads.Count() == 0)
                return;


            foreach (Dictionary<string, object> load in loads)
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

                        SqlDataReader results = ExecuteSQLReturnDataReader(DatabaseConnectionStringNames.PBSInvoices, CommandType.StoredProcedure, "Proc_Select_Header_Body_By_Bill_Date_No_Additional_Copies_By_CarrierID",
                                                                                        new SqlParameter("@psdatBillDate", load["bill_date"].ToString()),
                                                                                        new SqlParameter("@pvchrBillSource", load["bill_source"].ToString()),
                                                                                        new SqlParameter("@pvchrCarrier", carrier["carrier"].ToString()));

                        string outputFile = GetConfigurationKeyValue("PBSInvoicesByCarrierIdDirectory") + Convert.ToDateTime(load["bill_date"].ToString()).ToString("yyyy") + "\\";

                        //create the directory if it doesn't already exist
                        if (!Directory.Exists(outputFile))
                            Directory.CreateDirectory(outputFile);

                        outputFile += carrier["carrier"] + "_" + Convert.ToDateTime(load["bill_date"].ToString()).ToString("yyyyMMdd") + "_" + load["bill_source"].ToString() + ".pdf";

                        using (rptInvoices report = new rptInvoices())
                        {
                            report.SetDataSource((IDataReader)results);
                            report.ExportToDisk(ExportFormatType.PortableDocFormat, outputFile);
                        }

                        //run update sproc
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoices, "Proc_Update_Loads_Successful_AutoPrint_to_PDF_By_CarrierID_Flag",
                                                        new SqlParameter("@pintLoadsID", load["loads_id"].ToString()));

                        WriteToJobLog(JobLogMessageType.INFO, $"{outputFile} successfully created");

                    }
                }
            }

        }

        //This seems backwards from every other job, but the problem is that Crystal Reports seem to work best with data readers but the 
        // rest of the code works best (and is at this point the standard) with lists and dictionaries. I also don't want to make an extra
        // round trip to the database with results we already have stored in memory
        private Dictionary<string, object> ConvertDataReaderToDictionary(SqlDataReader sqlDataReader)
        {

            Dictionary<string, object> dictionary = new Dictionary<string, object>();

            for (int i = 0; i < sqlDataReader.FieldCount; i++)
            {
                dictionary.Add(sqlDataReader.GetName(i), sqlDataReader.GetValue(i));
            }
            return dictionary;

        }

        private void DeleteTemp()
        {
            //GDS - I'm not sure why this is needed here or why these temp files are getting created. From
            //the current logs, it looks like this code is still in use, so it has been carried over

            List<string> tempFiles = Directory.GetFiles(Path.GetTempPath(), "ctm*.tmp").ToList();

            foreach (string tempFile in tempFiles)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Delete temp file {tempFile}");
                File.Delete(tempFile);
            }

        }

    }
}
