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
                List<Dictionary<string, object>> configurationTables = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Configuration_Tables").ToList();


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
                                                                                        new SqlParameter("@pvchrBackupDir", GetConfigurationKeyValue("BackupDirectory")),
                                                                                        new SqlParameter("@pvchrBackupFile", fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".txt"),
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
                                                          new SqlParameter("@TypeOfBill", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString() ),
                                                          new SqlParameter("@AccountID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                                          new SqlParameter("@AreaCode", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                                          new SqlParameter("@BillDate", lineSegments[4].ToString() == "" ? (object)DBNull.Value : Convert.ToDateTime(lineSegments[4].ToString())),
                                                          new SqlParameter("@BillSourceID", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                                          new SqlParameter("@Building", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                                          new SqlParameter("@CityID", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                                          new SqlParameter("@CompanyID", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                                          new SqlParameter("@Complex", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                                                          new SqlParameter("@CountryID", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                                                          new SqlParameter("@CountyID", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                                          new SqlParameter("@CreditCardOnFile", lineSegments[12].ToString() == "no" ? 1 : 0),
                                                          new SqlParameter("@CurrentBillAmount", lineSegments[13].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[13].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                                          new SqlParameter("@DepotID", lineSegments[14].ToString() == "" ? (object)DBNull.Value : lineSegments[14].ToString()),
                                                          new SqlParameter("@DistributionCodeID", lineSegments[15].ToString() == "" ? (object)DBNull.Value : lineSegments[15].ToString()),
                                                          new SqlParameter("@DistrictID", lineSegments[16].ToString() == "" ? (object)DBNull.Value : lineSegments[16].ToString()),
                                                          new SqlParameter("@DropOrder", lineSegments[17].ToString() == "" ? (object)DBNull.Value : lineSegments[17].ToString()),
                                                          new SqlParameter("@DueDate", lineSegments[18].ToString() == "" ? (object)DBNull.Value : Convert.ToDateTime(lineSegments[18].ToString())),
                                                          new SqlParameter("@FirstName", lineSegments[19].ToString() == "" ? (object)DBNull.Value : lineSegments[19].ToString()),
                                                          new SqlParameter("@HonorificID", lineSegments[20].ToString() == "" ? (object)DBNull.Value : lineSegments[20].ToString()),
                                                          new SqlParameter("@HouseNumber", lineSegments[21].ToString() == "" ? (object)DBNull.Value : lineSegments[21].ToString()),
                                                          new SqlParameter("@HouseNumberModifier", lineSegments[22].ToString() == "" ? (object)DBNull.Value : lineSegments[22].ToString()),
                                                          new SqlParameter("@InvoiceNumber", lineSegments[23].ToString() == "" ? (object)DBNull.Value : Convert.ToInt32(lineSegments[23].ToString())),
                                                          new SqlParameter("@LastBillAmount", lineSegments[24].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[24].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                                          new SqlParameter("@LastBillDate", lineSegments[25].ToString() == "" ? (object)DBNull.Value : Convert.ToDateTime(lineSegments[25].ToString())),
                                                          new SqlParameter("@LastName", lineSegments[26].ToString() == "" ? (object)DBNull.Value : lineSegments[26].ToString()),
                                                          new SqlParameter("@MiddleInitial", lineSegments[27].ToString() == "" ? (object)DBNull.Value : lineSegments[27].ToString()),
                                                          new SqlParameter("@NameAddressLine1", lineSegments[28].ToString() == "" ? (object)DBNull.Value : lineSegments[28].ToString()),
                                                          new SqlParameter("@NameAddressLine2", lineSegments[29].ToString() == "" ? (object)DBNull.Value : lineSegments[29].ToString()),
                                                          new SqlParameter("@NameAddressLine3", lineSegments[30].ToString() == "" ? (object)DBNull.Value : lineSegments[30].ToString()),
                                                          new SqlParameter("@NameAddressLine4", lineSegments[31].ToString() == "" ? (object)DBNull.Value : lineSegments[31].ToString()),
                                                          new SqlParameter("@NameAddressLine5", lineSegments[32].ToString() == "" ? (object)DBNull.Value : lineSegments[32].ToString()),
                                                          new SqlParameter("@NameAddressLine6", lineSegments[33].ToString() == "" ? (object)DBNull.Value : lineSegments[33].ToString()),
                                                          new SqlParameter("@PastDueBalance", lineSegments[34].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[34].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                                          new SqlParameter("@Phone", lineSegments[35].ToString() == "" ? (object)DBNull.Value : lineSegments[35].ToString()),
                                                          new SqlParameter("@PostDirectional", lineSegments[36].ToString() == "" ? (object)DBNull.Value : lineSegments[36].ToString()),
                                                          new SqlParameter("@PreDirectional", lineSegments[37].ToString() == "" ? (object)DBNull.Value : lineSegments[37].ToString()),
                                                          new SqlParameter("@ProductID", lineSegments[38].ToString() == "" ? (object)DBNull.Value : lineSegments[38].ToString()),
                                                          new SqlParameter("@RemitToAddressLine1", lineSegments[39].ToString() == "" ? (object)DBNull.Value : lineSegments[39].ToString()),
                                                          new SqlParameter("@RemitToAddressLine2", lineSegments[40].ToString() == "" ? (object)DBNull.Value : lineSegments[40].ToString()),
                                                          new SqlParameter("@RemitToAddressLine3", lineSegments[41].ToString() == "" ? (object)DBNull.Value : lineSegments[41].ToString()),
                                                          new SqlParameter("@RemitToAddressLine4", lineSegments[42].ToString() == "" ? (object)DBNull.Value : lineSegments[42].ToString()),
                                                          new SqlParameter("@RemitToAddressLine5", lineSegments[43].ToString() == "" ? (object)DBNull.Value : lineSegments[43].ToString()),
                                                          new SqlParameter("@RemitToAddressLine6", lineSegments[44].ToString() == "" ? (object)DBNull.Value : lineSegments[44].ToString()),
                                                          new SqlParameter("@RemitToAddressLine7", lineSegments[45].ToString() == "" ? (object)DBNull.Value : lineSegments[45].ToString()),
                                                          new SqlParameter("@RouteID", lineSegments[46].ToString() == "" ? (object)DBNull.Value : lineSegments[46].ToString()),
                                                          new SqlParameter("@ScanLine", lineSegments[47].ToString() == "" ? (object)DBNull.Value : lineSegments[47].ToString()),
                                                          new SqlParameter("@StateID", lineSegments[48].ToString() == "" ? (object)DBNull.Value : lineSegments[48].ToString()),
                                                          new SqlParameter("@StreetName", lineSegments[49].ToString() == "" ? (object)DBNull.Value : lineSegments[49].ToString()),
                                                          new SqlParameter("@StreetSuffixID", lineSegments[50].ToString() == "" ? (object)DBNull.Value : lineSegments[50].ToString()),
                                                          new SqlParameter("@Terms", lineSegments[51].ToString() == "" ? (object)DBNull.Value : lineSegments[51].ToString()),
                                                          new SqlParameter("@TotalDue", lineSegments[52].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[52].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                                          new SqlParameter("@TruckID", lineSegments[53].ToString() == "" ? (object)DBNull.Value : lineSegments[53].ToString()),
                                                          new SqlParameter("@UnitDesignatorID", lineSegments[54].ToString() == "" ? (object)DBNull.Value : lineSegments[54].ToString()),
                                                          new SqlParameter("@UnitNumber", lineSegments[55].ToString() == "" ? (object)DBNull.Value : lineSegments[55].ToString()),
                                                          new SqlParameter("@ZipBarCode", lineSegments[56].ToString() == "" ? (object)DBNull.Value : lineSegments[56].ToString()),
                                                          new SqlParameter("@ZipCode", lineSegments[57].ToString() == "" ? (object)DBNull.Value : lineSegments[57].ToString()),
                                                          new SqlParameter("@ZipExtension", lineSegments[58].ToString() == "" ? (object)DBNull.Value : lineSegments[58].ToString()));
                    }
                    else if (lineSegments[0] == "Advance Draw Charge")
                    {
                        if (lineSegments[1] == "Bill")
                        {
                                 ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Advance_Draw_Charge_Bill",
                                                     new SqlParameter("@loads_id", loadsId),
                                                  new SqlParameter("@account_record_number", accountRecordNumber),
                                                  new SqlParameter("@load_sequence", lineNumber),
                                                  new SqlParameter("@AccountID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                                  new SqlParameter("@Amount", lineSegments[3].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[3].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                                  new SqlParameter("@BillSourceID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                                  new SqlParameter("@ChargeCodeID", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                                  new SqlParameter("@CompanyID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                                  new SqlParameter("@Description", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                                  new SqlParameter("@ProductID", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                                  new SqlParameter("@Quantity", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                                                  new SqlParameter("@RecapFormat", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                                                  new SqlParameter("@RecapID", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                                  new SqlParameter("@Reversal", lineSegments[12].ToString() == "" ? (object)DBNull.Value : lineSegments[12].ToString()),
                                                  new SqlParameter("@UnitRate", lineSegments[13].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[13].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)));
                        }
                    }
                    else if (lineSegments[0] == "Aging")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Aging",
                                             new SqlParameter("@loads_id", loadsId),
                                              new SqlParameter("@account_record_number", accountRecordNumber),
                                              new SqlParameter("@load_sequence", lineNumber),
                                              new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                              new SqlParameter("@AgeCurrent", lineSegments[2].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[2].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                              new SqlParameter("@AgePeriod1", lineSegments[3].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[3].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                              new SqlParameter("@AgePeriod2", lineSegments[4].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[4].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                              new SqlParameter("@AgePeriod3", lineSegments[5].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[5].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                              new SqlParameter("@AgePeriod4", lineSegments[6].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[6].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                              new SqlParameter("@BillSourceID", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                              new SqlParameter("@CompanyID", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()));
                    }
                    else if (lineSegments[0] == "Balance Forward")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Balance_Forward",
                                         new SqlParameter("@loads_id", loadsId),
                                          new SqlParameter("@account_record_number", accountRecordNumber),
                                          new SqlParameter("@load_sequence", lineNumber),
                                          new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                          new SqlParameter("@BillSourceID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                          new SqlParameter("@CompanyID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                          new SqlParameter("@LastBillAmount", lineSegments[4].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[4].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                          new SqlParameter("@LastBillDate", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()));
                    }
                    else if (lineSegments[0] == "Bill Message")
                    {
                    //    billMessageCount++;

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
                                      new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                      new SqlParameter("@Amount", lineSegments[2].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[2].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                      new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                      new SqlParameter("@CompanyID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                      new SqlParameter("@DueDate", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()));
                    }
                    else if (lineSegments[0] == "Debit Memo")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Debit_Memo",
                                         new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", accountRecordNumber),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                      new SqlParameter("@BillSourceID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                      new SqlParameter("@CompanyID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                      new SqlParameter("@Description", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                      new SqlParameter("@DueDate", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                      new SqlParameter("@OriginalAmount", lineSegments[6].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[6].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)));
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
                                          new SqlParameter("@AccountID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                          new SqlParameter("@Amount", lineSegments[3].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[3].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                          new SqlParameter("@BillSourceID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                          new SqlParameter("@ChargeCodeID", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                          new SqlParameter("@CompanyID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                          new SqlParameter("@Description", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                          new SqlParameter("@ProductID", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                          new SqlParameter("@Quantity", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                                          new SqlParameter("@RecapFormat", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                                          new SqlParameter("@RecapID", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                          new SqlParameter("@Reversal", lineSegments[12].ToString() == "no" ? 1 : 0),
                                          new SqlParameter("@UnitRate", lineSegments[13].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[13].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)));
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
                                      new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                      new SqlParameter("@Amount", lineSegments[2].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[2].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                      new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                      new SqlParameter("@ChargeCodeID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                      new SqlParameter("@ChargeDate", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                      new SqlParameter("@ChargeTypeID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                      new SqlParameter("@CompanyID", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                      new SqlParameter("@Description", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                      new SqlParameter("@ProductID", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                                      new SqlParameter("@Quantity", lineSegments[10].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[10].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                      new SqlParameter("@RecapFormat", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                      new SqlParameter("@RecapID", lineSegments[12].ToString() == "" ? (object)DBNull.Value : lineSegments[12].ToString()),
                                      new SqlParameter("@Remarks", lineSegments[13].ToString() == "" ? (object)DBNull.Value : lineSegments[13].ToString()),
                                      new SqlParameter("@RouteID", lineSegments[14].ToString() == "" ? (object)DBNull.Value : lineSegments[14].ToString()),
                                      new SqlParameter("@UnitRate", lineSegments[15].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[15].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)));
                    }
                    else if (lineSegments[0] == "Misc Charge")
                    {
                       ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Misc_Charge",
                                     new SqlParameter("@loads_id", loadsId),
                                  new SqlParameter("@account_record_number", accountRecordNumber),
                                  new SqlParameter("@load_sequence", lineNumber),
                                  new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                  new SqlParameter("@Amount", lineSegments[2].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[2].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                  new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                  new SqlParameter("@ChargeCodeID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                  new SqlParameter("@ChargeDate", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                  new SqlParameter("@ChargeTypeID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                  new SqlParameter("@CompanyID", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                  new SqlParameter("@Description", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                  new SqlParameter("@ProductID", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                                  new SqlParameter("@Quantity", lineSegments[10].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[10].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                  new SqlParameter("@RecapFormat", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                  new SqlParameter("@RecapID", lineSegments[12].ToString() == "" ? (object)DBNull.Value : lineSegments[12].ToString()),
                                  new SqlParameter("@Remarks", lineSegments[13].ToString() == "" ? (object)DBNull.Value : lineSegments[13].ToString()),
                                  new SqlParameter("@RouteID", lineSegments[14].ToString() == "" ? (object)DBNull.Value : lineSegments[14].ToString()),
                                  new SqlParameter("@UnitRate", lineSegments[15].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[15].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)));
                    }
                    else if (lineSegments[0] == "Misc Charge Reversal")
                    {
                       ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Misc_Charge_Reversal",
                                new SqlParameter("@loads_id", loadsId),
                              new SqlParameter("@account_record_number", accountRecordNumber),
                              new SqlParameter("@load_sequence", lineNumber),
                              new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                              new SqlParameter("@Amount", lineSegments[2].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[2].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                              new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                              new SqlParameter("@ChargeCodeID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                              new SqlParameter("@ChargeDate", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                              new SqlParameter("@ChargeTypeID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                              new SqlParameter("@CompanyID", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                              new SqlParameter("@Description", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                              new SqlParameter("@ProductID", lineSegments[9].ToString() == "" ? (object)DBNull.Value : lineSegments[9].ToString()),
                              new SqlParameter("@Quantity", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                              new SqlParameter("@RecapFormat", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                              new SqlParameter("@RecapID", lineSegments[12].ToString() == "" ? (object)DBNull.Value : lineSegments[12].ToString()),
                              new SqlParameter("@Remarks", lineSegments[13].ToString() == "" ? (object)DBNull.Value : lineSegments[13].ToString()),
                              new SqlParameter("@RouteID", lineSegments[14].ToString() == "" ? (object)DBNull.Value : lineSegments[14].ToString()),
                              new SqlParameter("@UnitRate", lineSegments[15].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[15].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)));
                    }
                    else if (lineSegments[0] == "Payment")
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Payment",
                                 new SqlParameter("@loads_id", loadsId),
                              new SqlParameter("@account_record_number", accountRecordNumber),
                              new SqlParameter("@load_sequence", lineNumber),
                              new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                              new SqlParameter("@Amount", lineSegments[2].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[2].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                              new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                              new SqlParameter("@CompanyID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                              new SqlParameter("@Remarks", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                              new SqlParameter("@TranDate", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()));
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
                                    new SqlParameter("@AccountID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                    new SqlParameter("@Amount", lineSegments[3].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[3].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                    new SqlParameter("@BillSourceID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                    new SqlParameter("@ChargeCodeID", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                    new SqlParameter("@CompanyID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                    new SqlParameter("@Description", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                    new SqlParameter("@ProductID", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                    new SqlParameter("@Quantity", lineSegments[9].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[9].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                    new SqlParameter("@RecapFormat", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                                    new SqlParameter("@RecapID", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                    new SqlParameter("@Reversal", lineSegments[12].ToString() == "no" ? 1 : 0),
                                    new SqlParameter("@UnitRate", lineSegments[13].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[13].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)));
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
                                  new SqlParameter("@AccountID", lineSegments[2].ToString() == "" ? (object)DBNull.Value : lineSegments[2].ToString()),
                                  new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                  new SqlParameter("@CompanyID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                  new SqlParameter("@DeliveryScheduleID", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()),
                                  new SqlParameter("@DistrictID", lineSegments[6].ToString() == "" ? (object)DBNull.Value : lineSegments[6].ToString()),
                                  new SqlParameter("@DrawClassID", lineSegments[7].ToString() == "" ? (object)DBNull.Value : lineSegments[7].ToString()),
                                  new SqlParameter("@DrawDate", lineSegments[8].ToString() == "" ? (object)DBNull.Value : lineSegments[8].ToString()),
                                  new SqlParameter("@DrawTotal", lineSegments[9].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[9].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                  new SqlParameter("@ProductID", lineSegments[10].ToString() == "" ? (object)DBNull.Value : lineSegments[10].ToString()),
                                  new SqlParameter("@RouteID", lineSegments[11].ToString() == "" ? (object)DBNull.Value : lineSegments[11].ToString()),
                                  new SqlParameter("@RouteType", lineSegments[12].ToString() == "" ? (object)DBNull.Value : lineSegments[12].ToString()),
                                  new SqlParameter("@SubstituteDelivery", lineSegments[12].ToString() == "no" ? 1 : 0));
                        }
                     }
                    else if (lineSegments[0] == "Total Due")
                    {
                        
                        ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Total_Due",
                                      new SqlParameter("@loads_id", loadsId),
                                      new SqlParameter("@account_record_number", accountRecordNumber),
                                      new SqlParameter("@load_sequence", lineNumber),
                                      new SqlParameter("@AccountID", lineSegments[1].ToString() == "" ? (object)DBNull.Value : lineSegments[1].ToString()),
                                      new SqlParameter("@Amount", lineSegments[2].ToString() == "" ? (object)DBNull.Value : Decimal.Parse(lineSegments[2].ToString(), System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowTrailingSign)),
                                      new SqlParameter("@BillSourceID", lineSegments[3].ToString() == "" ? (object)DBNull.Value : lineSegments[3].ToString()),
                                      new SqlParameter("@CompanyID", lineSegments[4].ToString() == "" ? (object)DBNull.Value : lineSegments[4].ToString()),
                                      new SqlParameter("@DueDate", lineSegments[5].ToString() == "" ? (object)DBNull.Value : lineSegments[5].ToString()));
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
