using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

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
                List<Dictionary<string, object>> configurationTables = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proce_Select_Configuration_Tables").ToList();


                //iterate and process files
                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                        if (previouslyLoadedFile == null)
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                            CopyAndProcessFile(fileInfo);
                        }
                        //else
                        //{
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.Wrappers, "Proc_Insert_Loads_Not_Loaded",
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
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Loads",
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


            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            DateTime? runDate = null;
            String runType = "";

            foreach (string line in fileContents)
            {
                int lineNumber = 0;
                int accountInformationCount = 0;
                int advanceDrawChargeBillCount = 0;
                int agingCount = 0;
                int balanceForwardCount = 0;
                int billMessageCount = 0;
                int collectionMessageCount = 0;
                int currentBillAmountCount = 0;
                int debitMemoCount = 0;
                int drawChargeBillCount = 0;
                int drawChargeDrawCount = 0;
                int dropCompensationCount = 0;
                int miscChargeCount = 0;
                int miscChargeReversalCount = 0;
                int paymentCount = 0;
                int returnsBillCount = 0;
                int returnsDrawCount = 0;
                int totalDueCount = 0;

                if (line != null && line.Trim().Length > 0)
                {
                    WriteToJobLog(JobLogMessageType.INFO, "Reading " + fileInfo.FullName);

                    lineNumber++;


                    List<string> lineSegments = line.Split('|').ToList();

                    if (lineSegments[0] == "Account Information")
                    {
                        accountInformationCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Account_Information",
                                                            new SqlParameter("@loads_id", loadsId),
                                                          new SqlParameter("@account_record_number", accountInformationCount),
                                                          new SqlParameter("@load_sequence", lineNumber),
                                                          new SqlParameter("@TypeOfBill", line[1]),
                                                          new SqlParameter("@AccountID", line[2]),
                                                          new SqlParameter("@AreaCode", line[3]),
                                                          new SqlParameter("@BillDate", line[4]),
                                                          new SqlParameter("@BillSourceID", line[5]),
                                                          new SqlParameter("@Building", line[6]),
                                                          new SqlParameter("@CityID", line[7]),
                                                          new SqlParameter("@CompanyID", line[8]),
                                                          new SqlParameter("@Complex", line[9]),
                                                          new SqlParameter("@CountryID", line[10]),
                                                          new SqlParameter("@CountyID", line[11]),
                                                          new SqlParameter("@CreditCardOnFile", line[12]),
                                                          new SqlParameter("@CurrentBillAmount", line[13]),
                                                          new SqlParameter("@DepotID", line[14]),
                                                          new SqlParameter("@DistributionCodeID", line[15]),
                                                          new SqlParameter("@DistrictID", line[16]),
                                                          new SqlParameter("@DropOrder", line[17]),
                                                          new SqlParameter("@DueDate", line[18]),
                                                          new SqlParameter("@FirstName", line[19]),
                                                          new SqlParameter("@HonorificID", line[20]),
                                                          new SqlParameter("@HouseNumber", line[21]),
                                                          new SqlParameter("@HouseNumberModifier", line[22]),
                                                          new SqlParameter("@InvoiceNumber", line[23]),
                                                          new SqlParameter("@LastBillAmount", line[24]),
                                                          new SqlParameter("@LastBillDate", line[25]),
                                                          new SqlParameter("@LastName", line[26]),
                                                          new SqlParameter("@MiddleInitial", line[27]),
                                                          new SqlParameter("@NameAddressLine1", line[28]),
                                                          new SqlParameter("@NameAddressLine2", line[29]),
                                                          new SqlParameter("@NameAddressLine3", line[30]),
                                                          new SqlParameter("@NameAddressLine4", line[31]),
                                                          new SqlParameter("@NameAddressLine5", line[32]),
                                                          new SqlParameter("@NameAddressLine6", line[33]),
                                                          new SqlParameter("@PastDueBalance", line[34]),
                                                          new SqlParameter("@Phone", line[35]),
                                                          new SqlParameter("@PostDirectional", line[36]),
                                                          new SqlParameter("@PreDirectional", line[37]),
                                                          new SqlParameter("@ProductID", line[38]),
                                                          new SqlParameter("@RemitToAddressLine1", line[39]),
                                                          new SqlParameter("@RemitToAddressLine2", line[40]),
                                                          new SqlParameter("@RemitToAddressLine3", line[41]),
                                                          new SqlParameter("@RemitToAddressLine4", line[42]),
                                                          new SqlParameter("@RemitToAddressLine5", line[43]),
                                                          new SqlParameter("@RemitToAddressLine6", line[44]),
                                                          new SqlParameter("@RemitToAddressLine7", line[45]),
                                                          new SqlParameter("@RouteID", line[46]),
                                                          new SqlParameter("@ScanLine", line[47]),
                                                          new SqlParameter("@StateID", line[48]),
                                                          new SqlParameter("@StreetName", line[49]),
                                                          new SqlParameter("@StreetSuffixID", line[50]),
                                                          new SqlParameter("@Terms", line[51]),
                                                          new SqlParameter("@TotalDue", line[52]),
                                                          new SqlParameter("@TruckID", line[53]),
                                                          new SqlParameter("@UnitDesignatorID", line[54]),
                                                          new SqlParameter("@UnitNumber", line[55]),
                                                          new SqlParameter("@ZipBarCode", line[56]),
                                                          new SqlParameter("@ZipCode", line[57]),
                                                          new SqlParameter("@ZipExtension", line[58]));
                    }
                    else if (lineSegments[0] == "Advance Draw Charge Bill")
                    {
                        advanceDrawChargeBillCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Advance_Draw_Charge_Bill",
                                                 new SqlParameter("@loads_id", loadsId),
                                              new SqlParameter("@account_record_number", advanceDrawChargeBillCount),
                                              new SqlParameter("@load_sequence", lineNumber),
                                              new SqlParameter("@AccountID", line[1]),
                                              new SqlParameter("@Amount", line[2]),
                                              new SqlParameter("@BillSourceID", line[3]),
                                              new SqlParameter("@ChargeCodeID", line[4]),
                                              new SqlParameter("@CompanyID", line[5]),
                                              new SqlParameter("@Description", line[6]),
                                              new SqlParameter("@ProductID", line[7]),
                                              new SqlParameter("@Quantity", line[8]),
                                              new SqlParameter("@RecapFormat", line[9]),
                                              new SqlParameter("@RecapID", line[10]),
                                              new SqlParameter("@Reversal", line[11]),
                                              new SqlParameter("@UnitRate", line[12]));
                    }
                    else if (lineSegments[0] == "Aging")
                    {
                        agingCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Aging",
                                             new SqlParameter("@loads_id", loadsId),
                                              new SqlParameter("@account_record_number", agingCount),
                                              new SqlParameter("@load_sequence", lineNumber),
                                              new SqlParameter("@AccountID", line[1]),
                                              new SqlParameter("@AgeCurrent", line[2]),
                                              new SqlParameter("@AgePeriod1", line[3]),
                                              new SqlParameter("@AgePeriod2", line[4]),
                                              new SqlParameter("@AgePeriod3", line[5]),
                                              new SqlParameter("@AgePeriod4", line[6]),
                                              new SqlParameter("@BillSourceID", line[7]),
                                              new SqlParameter("@CompanyID)", line[8]));
                    }
                    else if (lineSegments[0] == "Balance Forward")
                    {
                        balanceForwardCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Balance_Forward",
                                         new SqlParameter("@loads_id", loadsId),
                                          new SqlParameter("@account_record_number", balanceForwardCount),
                                          new SqlParameter("@load_sequence", lineNumber),
                                          new SqlParameter("@AccountID", line[1]),
                                          new SqlParameter("@BillSourceID", line[2]),
                                          new SqlParameter("@CompanyID", line[3]),
                                          new SqlParameter("@LastBillAmount", line[4]),
                                          new SqlParameter("@LastBillDate)", line[5]));
                    }
                    else if (lineSegments[0] == "Bill Message")
                    {
                        billMessageCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Bill_Message",
                                      new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", billMessageCount),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@BillSourceID", line[1]),
                                      new SqlParameter("@CompanyID", line[2]),
                                      new SqlParameter("@EntityID", line[3]),
                                      new SqlParameter("@EntityType", line[4]),
                                      new SqlParameter("@MessageText", line[5]),
                                      new SqlParameter("@PrintOrder", line[6]));
                    }
                    else if (lineSegments[0] == "Collection Message")
                    {
                        collectionMessageCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Collection_Message",
                                     new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", collectionMessageCount),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@BillSourceID", line[1]),
                                      new SqlParameter("@CollectMessage", line[2]),
                                      new SqlParameter("@CompanyID", line[3]));
                    }
                    else if (lineSegments[0] == "Current Bill Amount")
                    {
                        currentBillAmountCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Current_Bill_Amount",
                                      new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", currentBillAmountCount),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", line[1]),
                                      new SqlParameter("@Amount", line[2]),
                                      new SqlParameter("@BillSourceID", line[3]),
                                      new SqlParameter("@CompanyID", line[4]),
                                      new SqlParameter("@DueDate", line[5]));
                    }
                    else if (lineSegments[0] == "Debit Memo")
                    {
                        debitMemoCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Debit_Memo",
                                         new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", debitMemoCount),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", line[1]),
                                      new SqlParameter("@BillSourceID", line[2]),
                                      new SqlParameter("@CompanyID", line[3]),
                                      new SqlParameter("@Description", line[4]),
                                      new SqlParameter("@DueDate", line[5]),
                                      new SqlParameter("@OriginalAmount", line[6]));
                    }
                    else if (lineSegments[0] == "Draw Charge Bill")
                    {
                        drawChargeBillCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Draw_Charge_Bill",
                                              new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", drawChargeBillCount),
                                      new SqlParameter("@load_sequence", line[1]),
                                      new SqlParameter("@draw_group_number", line[2]),
                                      new SqlParameter("@AccountID", line[3]),
                                      new SqlParameter("@Amount", line[4]),
                                      new SqlParameter("@BillSourceID", line[5]),
                                      new SqlParameter("@ChargeCodeID", line[6]),
                                      new SqlParameter("@CompanyID", line[7]),
                                      new SqlParameter("@Description", line[8]),
                                      new SqlParameter("@ProductID", line[9]),
                                      new SqlParameter("@Quantity", line[10]),
                                      new SqlParameter("@RecapFormat", line[11]),
                                      new SqlParameter("@RecapID", line[12]),
                                      new SqlParameter("@Reversal", line[13]),
                                      new SqlParameter("@UnitRate", line[14]));
                    }
                    else if (lineSegments[0] == "Draw Chage Draw")
                    {
                        drawChargeDrawCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Draw_Charge_Draw",
                                         new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", drawChargeDrawCount),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@draw_group_number", line[1]),
                                      new SqlParameter("@AccountID", line[2]),
                                      new SqlParameter("@BillSourceID", line[3]),
                                      new SqlParameter("@CompanyID", line[4]),
                                      new SqlParameter("@DeliveryScheduleID", line[5]),
                                      new SqlParameter("@DistrictID", line[6]),
                                      new SqlParameter("@DrawClassID", line[7]),
                                      new SqlParameter("@DrawDate", line[8]),
                                      new SqlParameter("@DrawTotal", line[9]),
                                      new SqlParameter("@ProductID", line[10]),
                                      new SqlParameter("@RouteID", line[11]),
                                      new SqlParameter("@RouteType", line[12]),
                                      new SqlParameter("@SubstituteDelivery", line[13]));
                    }
                    else if (lineSegments[0] == "Drop Compensation")
                    {
                        dropCompensationCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Drop_Compensation",
                                            new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", dropCompensationCount),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", line[1]),
                                      new SqlParameter("@Amount", line[2]),
                                      new SqlParameter("@BillSourceID", line[3]),
                                      new SqlParameter("@ChargeCodeID", line[4]),
                                      new SqlParameter("@ChargeDate", line[5]),
                                      new SqlParameter("@ChargeTypeID", line[6]),
                                      new SqlParameter("@CompanyID", line[7]),
                                      new SqlParameter("@Description", line[8]),
                                      new SqlParameter("@ProductID", line[9]),
                                      new SqlParameter("@Quantity", line[10]),
                                      new SqlParameter("@RecapFormat", line[11]),
                                      new SqlParameter("@RecapID", line[12]),
                                      new SqlParameter("@Remarks", line[13]),
                                      new SqlParameter("@RouteID", line[14]),
                                      new SqlParameter("@UnitRate", line[1]));
                    }
                    else if (lineSegments[0] == "Misc Charge")
                    {
                        miscChargeCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Misc_Charge",
                                     new SqlParameter("@loads_id", loadsId),
                                  new SqlParameter("@account_record_number", miscChargeCount),
                                  new SqlParameter("@load_sequence", lineNumber),
                                  new SqlParameter("@AccountID", line[1]),
                                  new SqlParameter("@Amount", line[2]),
                                  new SqlParameter("@BillSourceID", line[3]),
                                  new SqlParameter("@ChargeCodeID", line[4]),
                                  new SqlParameter("@ChargeDate", line[5]),
                                  new SqlParameter("@ChargeTypeID", line[6]),
                                  new SqlParameter("@CompanyID", line[7]),
                                  new SqlParameter("@Description", line[8]),
                                  new SqlParameter("@ProductID", line[9]),
                                  new SqlParameter("@Quantity", line[10]),
                                  new SqlParameter("@RecapFormat", line[11]),
                                  new SqlParameter("@RecapID", line[12]),
                                  new SqlParameter("@Remarks", line[13]),
                                  new SqlParameter("@RouteID", line[14]),
                                  new SqlParameter("@UnitRate", line[15]));
                    }
                    else if (lineSegments[0] == "Misc Charge Reversal")
                    {
                        miscChargeReversalCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Misc_Charge_Reversal",
                                new SqlParameter("@loads_id", loadsId),
                              new SqlParameter("@account_record_number", miscChargeReversalCount),
                              new SqlParameter("@load_sequence", lineNumber),
                              new SqlParameter("@AccountID", line[1]),
                              new SqlParameter("@Amount", line[2]),
                              new SqlParameter("@BillSourceID", line[3]),
                              new SqlParameter("@ChargeCodeID", line[4]),
                              new SqlParameter("@ChargeDate", line[5]),
                              new SqlParameter("@ChargeTypeID", line[6]),
                              new SqlParameter("@CompanyID", line[7]),
                              new SqlParameter("@Description", line[8]),
                              new SqlParameter("@ProductID", line[9]),
                              new SqlParameter("@Quantity", line[10]),
                              new SqlParameter("@RecapFormat", line[11]),
                              new SqlParameter("@RecapID", line[12]),
                              new SqlParameter("@Remarks", line[13]),
                              new SqlParameter("@RouteID", line[14]),
                              new SqlParameter("@UnitRate", line[15]));
                    }
                    else if (lineSegments[0] == "Payment")
                    {
                        paymentCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Payment",
                                 new SqlParameter("@loads_id", loadsId),
                              new SqlParameter("@account_record_number", paymentCount),
                              new SqlParameter("@load_sequence", lineNumber),
                              new SqlParameter("@AccountID", line[1]),
                              new SqlParameter("@Amount", line[2]),
                              new SqlParameter("@BillSourceID", line[3]),
                              new SqlParameter("@CompanyID", line[4]),
                              new SqlParameter("@Remarks", line[5]),
                              new SqlParameter("@TranDate", line[6]));
                    }
                    else if (lineSegments[0] == "Returns Bill")
                    {
                        returnsBillCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Returns_Bill",
                                    new SqlParameter("@loads_id", loadsId),
                                new SqlParameter("@account_record_number", returnsBillCount),
                                new SqlParameter("@returns_group_number", lineNumber),
                                new SqlParameter("@load_sequence", line[1]),
                                new SqlParameter("@AccountID", line[2]),
                                new SqlParameter("@Amount", line[3]),
                                new SqlParameter("@BillSourceID", line[4]),
                                new SqlParameter("@ChargeCodeID", line[5]),
                                new SqlParameter("@CompanyID", line[6]),
                                new SqlParameter("@Description", line[7]),
                                new SqlParameter("@ProductID", line[8]),
                                new SqlParameter("@Quantity", line[9]),
                                new SqlParameter("@RecapFormat", line[10]),
                                new SqlParameter("@RecapID", line[11]),
                                new SqlParameter("@Reversal", line[12]),
                                new SqlParameter("@UnitRate", line[13]));
                     }
                    else if (lineSegments[0] == "Returns Draw")
                    {
                        returnsDrawCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Returns_Draw",
                                 new SqlParameter("@loads_id", loadsId),
                              new SqlParameter("@account_record_number", returnsDrawCount),
                              new SqlParameter("@load_sequence", lineNumber),
                              new SqlParameter("@returns_group_number", line[1]),
                              new SqlParameter("@AccountID", line[2]),
                              new SqlParameter("@BillSourceID", line[3]),
                              new SqlParameter("@CompanyID", line[4]),
                              new SqlParameter("@DeliveryScheduleID", line[5]),
                              new SqlParameter("@DistrictID", line[6]),
                              new SqlParameter("@DrawClassID", line[7]),
                              new SqlParameter("@DrawDate", line[8]),
                              new SqlParameter("@DrawTotal", line[9]),
                              new SqlParameter("@ProductID", line[10]),
                              new SqlParameter("@RouteID", line[11]),
                              new SqlParameter("@RouteType", line[12]),
                              new SqlParameter("@SubstituteDelivery)", line[13]));
                    }
                    else if (lineSegments[0] == "Total Due")
                    {
                        totalDueCount++;

                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Total_Due",
                                      new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", totalDueCount),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", line[1]),
                                      new SqlParameter("@Amount", line[2]),
                                      new SqlParameter("@BillSourceID", line[3]),
                                      new SqlParameter("@CompanyID", line[4]),
                                      new SqlParameter("@DueDate", line[5]));
                    }
                }

            }
        }
    }
}
