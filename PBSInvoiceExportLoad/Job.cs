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

namespace PBSInvoiceExportLoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "PBS Invoice Export Load";
            JobDescription = "Builds invoice export data";
            AppConfigSectionName = "PBSInvoiceExportLoad";
        }

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpcarrier.????????").ToList();

                files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpcwd.????????").ToList());
                files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpfree pub.????????").ToList());
                files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpnie.????????").ToList());
                files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpalb1.????????").ToList());


                //load configuration from configuration specific tables
                Dictionary<string, object> configurationGeneral = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Configuration_General").FirstOrDefault();  //there is only 1 entry in this table
                List<Dictionary<string, object>> loadFileConfigurations = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Configuration_Load_Files").ToList();
                List<Dictionary<string, object>> configurationTables = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Configuration_Tables").ToList();


                //iterate and process files
                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Length > 0) //ignore empty files
                        {

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Loads_If_Processed",
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
                                else
                                    WriteToJobLog(JobLogMessageType.INFO, "There's a chance the file is still getting updated, so we'll pick it up next run");

                            }
                            //else
                            //{
                            //    ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Loads_Not_Loaded",
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
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Loads",
                                                                                        new SqlParameter("@pvchrPBSGeneralServerInstance", GetConfigurationKeyValue("RemoteServerInstance")),
                                                                                        new SqlParameter("@pvchrPBSGeneralDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                                                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                                                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")),
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pvchrBackupDir", GetConfigurationKeyValue("BackupDirectory")),
                                                                                        new SqlParameter("@pvchrBackupFile", fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".txt"),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");


            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            DateTime? runDate = null;
            String runType = "";

            int lineNumber = 0;
            int accountRecordNumber = 0;

            int drawChargeDrawGroupCount = 0;
            string drawChargeDrawGroupProductId = "";
            int drawChargeBillGroupCount = 0;
            string drawChargeDrawBillProductId = "";

            int returnsBillGroupCount = 0;
            string returnsBillGroupProductId = "";
            int returnsDrawGroupCount = 0;
            string returnsDrawGroupProductId = "";

            WriteToJobLog(JobLogMessageType.INFO, "Reading " + fileInfo.FullName);

            foreach (string line in fileContents)
            {

                if (line != null && line.Trim().Length > 0)
                {

                    lineNumber++;


                    List<string> lineSegments = line.Split('|').ToList();

                    if (lineSegments[0] == "Account Information")
                    {
                        if (lineSegments[5].ToString() != "")
                            runType = lineSegments[5].ToString();

                        if (lineSegments[4].ToString() != "")
                            runDate = Convert.ToDateTime(lineSegments[4].ToString());


                        accountRecordNumber++;
                        
                        drawChargeDrawGroupCount = 0;
                        drawChargeBillGroupCount = 0;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Account_Information",
                                                            new SqlParameter("@loads_id", loadsId),
                                                          new SqlParameter("@account_record_number", accountRecordNumber),
                                                          new SqlParameter("@load_sequence", lineNumber),
                                                          new SqlParameter("@TypeOfBill", FormatString(lineSegments[1].ToString())),
                                                          new SqlParameter("@AccountID", FormatString(lineSegments[2].ToString())),
                                                          new SqlParameter("@AreaCode", FormatString(lineSegments[3].ToString())),
                                                          new SqlParameter("@BillDate", FormatDateTime(lineSegments[4].ToString())),
                                                          new SqlParameter("@BillSourceID", FormatString(lineSegments[5].ToString())),
                                                          new SqlParameter("@Building", FormatString(lineSegments[6].ToString())),
                                                          new SqlParameter("@CityID", FormatString(lineSegments[7].ToString())),
                                                          new SqlParameter("@CompanyID", FormatString(lineSegments[8].ToString())),
                                                          new SqlParameter("@Complex", FormatString(lineSegments[9].ToString())),
                                                          new SqlParameter("@CountryID", FormatString(lineSegments[10].ToString())),
                                                          new SqlParameter("@CountyID", FormatString(lineSegments[11].ToString())),
                                                          new SqlParameter("@CreditCardOnFile", lineSegments[12].ToString() == "no" ? 1 : 0),
                                                          new SqlParameter("@CurrentBillAmount", FormatNumber(lineSegments[13].ToString())),
                                                          new SqlParameter("@DepotID", FormatString(lineSegments[14].ToString())),
                                                          new SqlParameter("@DistributionCodeID", FormatString(lineSegments[15].ToString())),
                                                          new SqlParameter("@DistrictID", FormatString(lineSegments[16].ToString())),
                                                          new SqlParameter("@DropOrder", FormatString(lineSegments[17].ToString())),
                                                          new SqlParameter("@DueDate", FormatString(lineSegments[18].ToString())),
                                                          new SqlParameter("@FirstName", FormatString(lineSegments[19].ToString())),
                                                          new SqlParameter("@HonorificID", FormatString(lineSegments[20].ToString())),
                                                          new SqlParameter("@HouseNumber", FormatString(lineSegments[21].ToString())),
                                                          new SqlParameter("@HouseNumberModifier", FormatString(lineSegments[22].ToString())),
                                                          new SqlParameter("@InvoiceNumber", FormatString(lineSegments[23].ToString())),
                                                          new SqlParameter("@LastBillAmount", FormatNumber(lineSegments[24].ToString())),
                                                          new SqlParameter("@LastBillDate", FormatString(lineSegments[25].ToString())),
                                                          new SqlParameter("@LastName", FormatString(lineSegments[26].ToString())),
                                                          new SqlParameter("@MiddleInitial", FormatString(lineSegments[27].ToString())),
                                                          new SqlParameter("@NameAddressLine1", FormatString(lineSegments[28].ToString())),
                                                          new SqlParameter("@NameAddressLine2", FormatString(lineSegments[29].ToString())),
                                                          new SqlParameter("@NameAddressLine3", FormatString(lineSegments[30].ToString())),
                                                          new SqlParameter("@NameAddressLine4", FormatString(lineSegments[31].ToString())),
                                                          new SqlParameter("@NameAddressLine5", FormatString(lineSegments[32].ToString())),
                                                          new SqlParameter("@NameAddressLine6", FormatString(lineSegments[33].ToString())),
                                                          new SqlParameter("@PastDueBalance", FormatNumber(lineSegments[34].ToString())),
                                                          new SqlParameter("@Phone", FormatString(lineSegments[35].ToString())),
                                                          new SqlParameter("@PostDirectional", FormatString(lineSegments[36].ToString())),
                                                          new SqlParameter("@PreDirectional", FormatString(lineSegments[37].ToString())),
                                                          new SqlParameter("@ProductID", FormatString(lineSegments[38].ToString())),
                                                          new SqlParameter("@RemitToAddressLine1", FormatString(lineSegments[39].ToString())),
                                                          new SqlParameter("@RemitToAddressLine2", FormatString(lineSegments[40].ToString())),
                                                          new SqlParameter("@RemitToAddressLine3", FormatString(lineSegments[41].ToString())),
                                                          new SqlParameter("@RemitToAddressLine4", FormatString(lineSegments[42].ToString())),
                                                          new SqlParameter("@RemitToAddressLine5", FormatString(lineSegments[43].ToString())),
                                                          new SqlParameter("@RemitToAddressLine6", FormatString(lineSegments[44].ToString())),
                                                          new SqlParameter("@RemitToAddressLine7", FormatString(lineSegments[45].ToString())),
                                                          new SqlParameter("@RouteID", FormatString(lineSegments[46].ToString())),
                                                          new SqlParameter("@ScanLine", FormatString(lineSegments[47].ToString())),
                                                          new SqlParameter("@StateID", FormatString(lineSegments[48].ToString())),
                                                          new SqlParameter("@StreetName", FormatString(lineSegments[49].ToString())),
                                                          new SqlParameter("@StreetSuffixID", FormatString(lineSegments[50].ToString())),
                                                          new SqlParameter("@Terms", FormatString(lineSegments[51].ToString())),
                                                          new SqlParameter("@TotalDue", FormatNumber(lineSegments[52].ToString())),
                                                          new SqlParameter("@TruckID", FormatString(lineSegments[53].ToString())),
                                                          new SqlParameter("@UnitDesignatorID", FormatString(lineSegments[54].ToString())),
                                                          new SqlParameter("@UnitNumber", FormatString(lineSegments[55].ToString())),
                                                          new SqlParameter("@ZipBarCode", FormatString(lineSegments[56].ToString())),
                                                          new SqlParameter("@ZipCode", FormatString(lineSegments[57].ToString())),
                                                          new SqlParameter("@ZipExtension", FormatString(lineSegments[58].ToString())));
                    }
                    else if (lineSegments[0] == "Advance Draw Charge")
                    {
                        if (lineSegments[1] == "Bill")
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Advance_Draw_Charge_Bill",
                                                new SqlParameter("@loads_id", loadsId),
                                             new SqlParameter("@account_record_number", accountRecordNumber),
                                             new SqlParameter("@load_sequence", lineNumber),
                                             new SqlParameter("@AccountID", FormatString(lineSegments[2].ToString())),
                                             new SqlParameter("@Amount", FormatNumber(lineSegments[3].ToString())),
                                             new SqlParameter("@BillSourceID", FormatString(lineSegments[4].ToString())),
                                             new SqlParameter("@ChargeCodeID", FormatString(lineSegments[5].ToString())),
                                             new SqlParameter("@CompanyID", FormatString(lineSegments[6].ToString())),
                                             new SqlParameter("@Description", FormatString(lineSegments[7].ToString())),
                                             new SqlParameter("@ProductID", FormatString(lineSegments[8].ToString())),
                                             new SqlParameter("@Quantity", FormatString(lineSegments[9].ToString())),
                                             new SqlParameter("@RecapFormat", FormatString(lineSegments[10].ToString())),
                                             new SqlParameter("@RecapID", FormatString(lineSegments[11].ToString())),
                                             new SqlParameter("@Reversal", FormatString(lineSegments[12].ToString())),
                                             new SqlParameter("@UnitRate", FormatNumber(lineSegments[13].ToString())));
                        }
                    }
                    else if (lineSegments[0] == "Aging")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Aging",
                                             new SqlParameter("@loads_id", loadsId),
                                              new SqlParameter("@account_record_number", accountRecordNumber),
                                              new SqlParameter("@load_sequence", lineNumber),
                                              new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                                              new SqlParameter("@AgeCurrent", FormatNumber(lineSegments[2].ToString())),
                                              new SqlParameter("@AgePeriod1", FormatNumber(lineSegments[3].ToString())),
                                              new SqlParameter("@AgePeriod2", FormatNumber(lineSegments[4].ToString())),
                                              new SqlParameter("@AgePeriod3", FormatNumber(lineSegments[5].ToString())),
                                              new SqlParameter("@AgePeriod4", FormatNumber(lineSegments[6].ToString())),
                                              new SqlParameter("@BillSourceID", FormatString(lineSegments[7].ToString())),
                                              new SqlParameter("@CompanyID", FormatString(lineSegments[8].ToString())));
                    }
                    else if (lineSegments[0] == "Balance Forward")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Balance_Forward",
                                         new SqlParameter("@loads_id", loadsId),
                                          new SqlParameter("@account_record_number", accountRecordNumber),
                                          new SqlParameter("@load_sequence", lineNumber),
                                          new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                                          new SqlParameter("@BillSourceID", FormatString(lineSegments[2].ToString())),
                                          new SqlParameter("@CompanyID", FormatString(lineSegments[3].ToString())),
                                          new SqlParameter("@LastBillAmount", FormatNumber(lineSegments[4].ToString())),
                                          new SqlParameter("@LastBillDate", FormatDateTime(lineSegments[5].ToString())));
                    }
                    else if (lineSegments[0] == "Bill Message")
                    {

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Bill_Message",
                                      new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", accountRecordNumber),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@BillSourceID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                      new SqlParameter("@CompanyID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                      new SqlParameter("@EntityID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                      new SqlParameter("@EntityType", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                      new SqlParameter("@MessageText", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                      new SqlParameter("@PrintOrder", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()));
                    }
                    else if (lineSegments[0] == "Collection Message")
                    {
                       ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Collection_Message",
                                     new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", accountRecordNumber),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@BillSourceID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                      new SqlParameter("@CollectMessage", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                      new SqlParameter("@CompanyID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()));
                    }
                    else if (lineSegments[0] == "Current Bill Amount")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Current_Bill_Amount",
                                       new SqlParameter("@loads_id", loadsId),
                                       new SqlParameter("@account_record_number", accountRecordNumber),
                                       new SqlParameter("@load_sequence", lineNumber),
                                       new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                                       new SqlParameter("@Amount", FormatNumber(lineSegments[2].ToString())),
                                       new SqlParameter("@BillSourceID", FormatString(lineSegments[3].ToString())),
                                       new SqlParameter("@CompanyID", FormatString(lineSegments[4].ToString())),
                                       new SqlParameter("@DueDate", FormatDateTime(lineSegments[5].ToString())));
                    }
                    else if (lineSegments[0] == "Debit Memo")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Debit_Memo",
                                         new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", accountRecordNumber),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                                      new SqlParameter("@BillSourceID", FormatString(lineSegments[2].ToString())),
                                      new SqlParameter("@CompanyID", FormatString(lineSegments[3].ToString())),
                                      new SqlParameter("@Description", FormatString(lineSegments[4].ToString())),
                                      new SqlParameter("@DueDate", FormatDateTime(lineSegments[5].ToString())),
                                      new SqlParameter("@OriginalAmount", FormatNumber(lineSegments[6].ToString())));
                    }
                    else if (lineSegments[0] == "Draw Charge")
                    {
                        if (lineSegments[1] == "Bill")
                        {

                            if (drawChargeBillGroupCount == 0)
                                drawChargeBillGroupCount++;
                            else if (drawChargeDrawBillProductId != lineSegments[11].ToString())
                            {
                                drawChargeBillGroupCount++;
                                drawChargeDrawBillProductId = lineSegments[11].ToString();
                            }

                            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Draw_Charge_Bill",
                                                  new SqlParameter("@loads_id", loadsId),
                                          new SqlParameter("@account_record_number", accountRecordNumber),
                                          new SqlParameter("@load_sequence", lineNumber),
                                          new SqlParameter("@draw_group_number", drawChargeBillGroupCount),
                                          new SqlParameter("@AccountID", FormatString(lineSegments[2].ToString())),
                                          new SqlParameter("@Amount", FormatNumber(lineSegments[3].ToString())),
                                          new SqlParameter("@BillSourceID", FormatString(lineSegments[4].ToString())),
                                          new SqlParameter("@ChargeCodeID", FormatString(lineSegments[5].ToString())),
                                          new SqlParameter("@CompanyID", FormatString(lineSegments[6].ToString())),
                                          new SqlParameter("@Description", FormatString(lineSegments[7].ToString())),
                                          new SqlParameter("@ProductID", FormatString(lineSegments[8].ToString())),
                                          new SqlParameter("@Quantity", FormatString(lineSegments[9].ToString())),
                                          new SqlParameter("@RecapFormat", FormatString(lineSegments[10].ToString())),
                                          new SqlParameter("@RecapID", FormatString(lineSegments[11].ToString())),
                                          new SqlParameter("@Reversal", lineSegments[12].ToString() == "no" ? 1 : 0),
                                          new SqlParameter("@UnitRate", FormatNumber(lineSegments[13].ToString())));
                        }
                        else if (lineSegments[1] == "Draw")
                        {
                         //   drawChargeDrawCount++;

                            if (drawChargeDrawGroupCount == 0)
                                drawChargeDrawGroupCount++;
                            else if (drawChargeDrawGroupProductId != lineSegments[11].ToString())
                            {
                                drawChargeDrawGroupCount++;
                                drawChargeDrawGroupProductId = lineSegments[11].ToString();
                            }

                            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Draw_Charge_Draw",
                                             new SqlParameter("@loads_id", loadsId),
                                          new SqlParameter("@account_record_number", accountRecordNumber),
                                          new SqlParameter("@load_sequence", lineNumber),
                                          new SqlParameter("@draw_group_number", drawChargeDrawGroupCount),
                                          new SqlParameter("@AccountID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                          new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                          new SqlParameter("@CompanyID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                          new SqlParameter("@DeliveryScheduleID", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                          new SqlParameter("@DistrictID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                          new SqlParameter("@DrawClassID", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                          new SqlParameter("@DrawDate", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                          new SqlParameter("@DrawTotal", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                                          new SqlParameter("@ProductID", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                                          new SqlParameter("@RouteID", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                          new SqlParameter("@RouteType", lineSegments[12].ToString() == "" ? (object)DBNull.Value : lineSegments[12].ToString()),
                                          new SqlParameter("@SubstituteDelivery", lineSegments[13].ToString() == "no" ? 1 : 0));
                        }
                    }
                    else if (lineSegments[0] == "Drop Compensation")
                    {

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Drop_Compensation",
                                            new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", accountRecordNumber),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                                      new SqlParameter("@Amount", FormatNumber(lineSegments[2].ToString())),
                                      new SqlParameter("@BillSourceID", FormatString(lineSegments[3].ToString())),
                                      new SqlParameter("@ChargeCodeID", FormatString(lineSegments[4].ToString())),
                                      new SqlParameter("@ChargeDate", FormatString(lineSegments[5].ToString())),
                                      new SqlParameter("@ChargeTypeID", FormatString(lineSegments[6].ToString())),
                                      new SqlParameter("@CompanyID", FormatString(lineSegments[7].ToString())),
                                      new SqlParameter("@Description", FormatString(lineSegments[8].ToString())),
                                      new SqlParameter("@ProductID", FormatString(lineSegments[9].ToString())),
                                      new SqlParameter("@Quantity", FormatNumber(lineSegments[10].ToString())),
                                      new SqlParameter("@RecapFormat", FormatString(lineSegments[11].ToString())),
                                      new SqlParameter("@RecapID", FormatString(lineSegments[12].ToString())),
                                      new SqlParameter("@Remarks", FormatString(lineSegments[13].ToString())),
                                      new SqlParameter("@RouteID", FormatString(lineSegments[14].ToString())),
                                      new SqlParameter("@UnitRate", FormatNumber(lineSegments[15].ToString())));
                    }
                    else if (lineSegments[0] == "Misc Charge")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Misc_Charge",
                                      new SqlParameter("@loads_id", loadsId),
                                   new SqlParameter("@account_record_number", accountRecordNumber),
                                   new SqlParameter("@load_sequence", lineNumber),
                                   new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                                   new SqlParameter("@Amount", FormatNumber(lineSegments[2].ToString())),
                                   new SqlParameter("@BillSourceID", FormatString(lineSegments[3].ToString())),
                                   new SqlParameter("@ChargeCodeID", FormatString(lineSegments[4].ToString())),
                                   new SqlParameter("@ChargeDate", FormatDateTime(lineSegments[5].ToString())),
                                   new SqlParameter("@ChargeTypeID", FormatString(lineSegments[6].ToString())),
                                   new SqlParameter("@CompanyID", FormatString(lineSegments[7].ToString())),
                                   new SqlParameter("@Description", FormatString(lineSegments[8].ToString())),
                                   new SqlParameter("@ProductID", FormatString(lineSegments[9].ToString())),
                                   new SqlParameter("@Quantity", FormatNumber(lineSegments[10].ToString())),
                                   new SqlParameter("@RecapFormat", FormatString(lineSegments[11].ToString())),
                                   new SqlParameter("@RecapID", FormatString(lineSegments[12].ToString())),
                                   new SqlParameter("@Remarks", FormatString(lineSegments[13].ToString())),
                                   new SqlParameter("@RouteID", FormatString(lineSegments[14].ToString())),
                                   new SqlParameter("@UnitRate", FormatNumber(lineSegments[15].ToString())));
                    }
                    else if (lineSegments[0] == "Misc Charge Reversal")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Misc_Charge_Reversal",
                                 new SqlParameter("@loads_id", loadsId),
                               new SqlParameter("@account_record_number", accountRecordNumber),
                               new SqlParameter("@load_sequence", lineNumber),
                               new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                               new SqlParameter("@Amount", FormatNumber(lineSegments[2].ToString())),
                               new SqlParameter("@BillSourceID", FormatString(lineSegments[3].ToString())),
                               new SqlParameter("@ChargeCodeID", FormatString(lineSegments[4].ToString())),
                               new SqlParameter("@ChargeDate", FormatString(lineSegments[5].ToString())),
                               new SqlParameter("@ChargeTypeID", FormatString(lineSegments[6].ToString())),
                               new SqlParameter("@CompanyID", FormatString(lineSegments[7].ToString())),
                               new SqlParameter("@Description", FormatString(lineSegments[8].ToString())),
                               new SqlParameter("@ProductID", FormatString(lineSegments[9].ToString())),
                               new SqlParameter("@Quantity", FormatNumber(lineSegments[10].ToString())),
                               new SqlParameter("@RecapFormat", FormatString(lineSegments[11].ToString())),
                               new SqlParameter("@RecapID", FormatString(lineSegments[12].ToString())),
                               new SqlParameter("@Remarks", FormatString(lineSegments[13].ToString())),
                               new SqlParameter("@RouteID", FormatString(lineSegments[14].ToString())),
                               new SqlParameter("@UnitRate", FormatNumber(lineSegments[15].ToString())));
                    }
                    else if (lineSegments[0] == "Payment")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Payment",
                                 new SqlParameter("@loads_id", loadsId),
                              new SqlParameter("@account_record_number", accountRecordNumber),
                              new SqlParameter("@load_sequence", lineNumber),
                              new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                              new SqlParameter("@Amount", FormatNumber(lineSegments[2].ToString())),
                              new SqlParameter("@BillSourceID", FormatString(lineSegments[3].ToString())),
                              new SqlParameter("@CompanyID", FormatString(lineSegments[4].ToString())),
                              new SqlParameter("@Remarks", FormatString(lineSegments[5].ToString())),
                              new SqlParameter("@TranDate", FormatDateTime(lineSegments[6].ToString())));
                    }
                    else if (lineSegments[0] == "Returns")
                    {
                        if (lineSegments[1] == "Bill")
                        {
                            if (returnsBillGroupCount == 0)
                                returnsBillGroupCount++;
                            else if (returnsBillGroupProductId != lineSegments[8].ToString())
                            {
                                returnsBillGroupCount++;
                                returnsBillGroupProductId = lineSegments[8].ToString();
                            }

                            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Returns_Bill",
                                        new SqlParameter("@loads_id", loadsId),
                                    new SqlParameter("@account_record_number", accountRecordNumber),
                                    new SqlParameter("@returns_group_number", returnsBillGroupCount),
                                    new SqlParameter("@load_sequence", lineNumber),
                                    new SqlParameter("@AccountID", FormatString(lineSegments[2].ToString())),
                                    new SqlParameter("@Amount", FormatNumber(lineSegments[3].ToString())),
                                    new SqlParameter("@BillSourceID", FormatString(lineSegments[4].ToString())),
                                    new SqlParameter("@ChargeCodeID", FormatString(lineSegments[5].ToString())),
                                    new SqlParameter("@CompanyID", FormatString(lineSegments[6].ToString())),
                                    new SqlParameter("@Description", FormatString(lineSegments[7].ToString())),
                                    new SqlParameter("@ProductID", FormatString(lineSegments[8].ToString())),
                                    new SqlParameter("@Quantity", FormatNumber(lineSegments[9].ToString())),
                                    new SqlParameter("@RecapFormat", FormatString(lineSegments[10].ToString())),
                                    new SqlParameter("@RecapID", FormatString(lineSegments[11].ToString())),
                                    new SqlParameter("@Reversal", lineSegments[12].ToString() == "no" ? 1 : 0),
                                    new SqlParameter("@UnitRate", FormatNumber(lineSegments[13].ToString())));
                        }
                        else if (lineSegments[1] == "Draw")
                        {
                            if (returnsDrawGroupCount == 0)
                                returnsDrawGroupCount++;
                            else if (returnsDrawGroupProductId != lineSegments[10].ToString())
                            {
                                returnsDrawGroupCount++;
                                returnsDrawGroupProductId = lineSegments[10].ToString();
                            }

                            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Returns_Draw",
                                     new SqlParameter("@loads_id", loadsId),
                                  new SqlParameter("@account_record_number", accountRecordNumber),
                                  new SqlParameter("@load_sequence", lineNumber),
                                  new SqlParameter("@returns_group_number", returnsDrawGroupCount),
                                  new SqlParameter("@AccountID", FormatString(lineSegments[2].ToString())),
                                  new SqlParameter("@BillSourceID", FormatString(lineSegments[3].ToString())),
                                  new SqlParameter("@CompanyID", FormatString(lineSegments[4].ToString())),
                                  new SqlParameter("@DeliveryScheduleID", FormatString(lineSegments[5].ToString())),
                                  new SqlParameter("@DistrictID", FormatString(lineSegments[6].ToString())),
                                  new SqlParameter("@DrawClassID", FormatString(lineSegments[7].ToString())),
                                  new SqlParameter("@DrawDate", FormatDateTime(lineSegments[8].ToString())),
                                  new SqlParameter("@DrawTotal", FormatNumber(lineSegments[9].ToString())),
                                  new SqlParameter("@ProductID", FormatString(lineSegments[10].ToString())),
                                  new SqlParameter("@RouteID", FormatString(lineSegments[11].ToString())),
                                  new SqlParameter("@RouteType", FormatString(lineSegments[12].ToString())),
                                  new SqlParameter("@SubstituteDelivery", lineSegments[13].ToString() == "no" ? 1 : 0));
                        }
                     }
                    else if (lineSegments[0] == "Total Due")
                    {

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Total_Due",
                                      new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", accountRecordNumber),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", FormatString(lineSegments[1].ToString())),
                                      new SqlParameter("@Amount", FormatNumber(lineSegments[2].ToString())),
                                      new SqlParameter("@BillSourceID", FormatString(lineSegments[3].ToString())),
                                      new SqlParameter("@CompanyID", FormatString(lineSegments[4].ToString())),
                                      new SqlParameter("@DueDate", FormatDateTime(lineSegments[5].ToString())));
                    }
                }

            }

            WriteToJobLog(JobLogMessageType.INFO, $"{lineNumber} total records read.");
            ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Update_Loads",
                                     new SqlParameter("@pintLoadsID", loadsId),
                                new SqlParameter("@pvchrBillSourceID", runType),
                                new SqlParameter("@pvchrBillDate", runDate.Value.ToShortDateString()),
                                new SqlParameter("@pintRecordCount", lineNumber),
                                new SqlParameter("@pflgSuccessfulLoad", true));
            WriteToJobLog(JobLogMessageType.INFO, "Load information updated");
        }
    }
}
