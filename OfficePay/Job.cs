using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OfficePay
{
    public class Job : JobBase
    {

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "renewals??????-???").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Select_Loads_If_Processed",
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
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "Proc_Insert_Loads_Not_Loaded",
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
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName, true);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.OfficePay, "Proc_Insert_Loads",
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

            ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            Int32 lineNumber = 0;
            Int32 processedSegmentCount = 0;

            foreach (string line in fileContents)
            {
                lineNumber++;

                if (line != null && line.Trim().Length > 0)
                {
                    List<string> lineSegments = line.Split('|').ToList();
                    List<string> segmentNames = new List<string>() { "F1", "B1", "C1", "D1", "X1", "EG", "E1", "G1", "G2", "G3",
                                                                     "G4", "Z0", "Z1", "Z2", "M1", "M2", "P1", "P2", "R1", "R2",
                                                                     "SG", "S1", "S2", "T1", "TC" };

                    //since one line can have multiple instances of the same segment, we need a way to keep track of each
                    Int32 F1Count = 0;
                    Int32 B1Count = 0;
                    Int32 C1Count = 0;
                    Int32 D1Count = 0;
                    Int32 X1Count = 0;
                    Int32 EGCount = 0;
                    Int32 E1Count = 0;
                    Int32 G1Count = 0;
                    Int32 G2Count = 0;
                    Int32 G3Count = 0;
                    Int32 G4Count = 0;
                    Int32 Z0Count = 0;
                    Int32 Z1Count = 0;
                    Int32 Z2Count = 0;
                    Int32 M1Count = 0;
                    Int32 M2Count = 0;
                    Int32 P1Count = 0;
                    Int32 P2Count = 0;
                    Int32 R1Count = 0;
                    Int32 R2Count = 0;
                    Int32 SGCount = 0;
                    Int32 S1Count = 0;
                    Int32 S2Count = 0;
                    Int32 T1Count = 0;
                    Int32 TCCount = 0;


                    string currentSegmentName = lineSegments[0];
                    Int32 lineSegmentCounter = 0;

                    foreach (string lineSegment in lineSegments)
                    {


                        //check to see if we are starting a new segment
                        if (segmentNames.Where(s => s == lineSegment).FirstOrDefault() != null)
                        {
                            currentSegmentName = lineSegment;

                            switch (currentSegmentName)
                            {
                                case "SG":
                                    SGCount++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Start_SG",
                                                   new SqlParameter("@loads_id", loadsId),
                                                   new SqlParameter("@pbs_record_number", lineNumber),
                                                   new SqlParameter("@segment_instance", SGCount),
                                                   new SqlParameter("@bill_to_name", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                   new SqlParameter("@bill_to_address_1", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                   new SqlParameter("@bill_to_address_2", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                   new SqlParameter("@bill_to_address_3", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                   new SqlParameter("@zip", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                   new SqlParameter("@zip_extension", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                                   new SqlParameter("@delivery_point_code", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                                   new SqlParameter("@route", FormatString(lineSegments[lineSegmentCounter + 8].ToString())),
                                                   new SqlParameter("@imb", FormatString(lineSegments[lineSegmentCounter + 9].ToString())),
                                                   new SqlParameter("@encoded_imb_mixed_case", FormatString(lineSegments[lineSegmentCounter + 10].ToString())),
                                                   new SqlParameter("@encoded_imb_uppercase", FormatString(lineSegments[lineSegmentCounter + 11].ToString())),
                                                   new SqlParameter("@last_field", (object)DBNull.Value));
                                    break;
                                case "S1":
                                    S1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Subscription_S1",
                                               new SqlParameter("@loads_id", loadsId),
                                               new SqlParameter("@pbs_record_number", lineNumber),
                                               new SqlParameter("@segment_instance", S1Count),
                                               new SqlParameter("@subscription_number", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                               new SqlParameter("@delivery_schedule", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                               new SqlParameter("@expire_date", FormatDateTime(lineSegments[lineSegmentCounter + 3].ToString())),
                                               new SqlParameter("@end_grace_date", FormatDateTime(lineSegments[lineSegmentCounter + 4].ToString())),
                                               new SqlParameter("@days_of_week_1", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                               new SqlParameter("@days_of_week_2", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                               new SqlParameter("@publication_name", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                               new SqlParameter("@lockbox_scanline_data", FormatString(lineSegments[lineSegmentCounter + 8].ToString())),
                                               new SqlParameter("@renewal_run_date", FormatDateTime(lineSegments[lineSegmentCounter + 9].ToString())),
                                               new SqlParameter("@renewal_invoice_or_grace", FormatString(lineSegments[lineSegmentCounter + 10].ToString())),
                                               new SqlParameter("@last_renewal_date", FormatDateTime(lineSegments[lineSegmentCounter + 11].ToString())),
                                               new SqlParameter("@renewal_number", FormatString(lineSegments[lineSegmentCounter + 12].ToString())),
                                               new SqlParameter("@start_reason_code", FormatString(lineSegments[lineSegmentCounter + 13].ToString())),
                                               new SqlParameter("@number_of_payments_since_start", FormatString(lineSegments[lineSegmentCounter + 14].ToString())),
                                               new SqlParameter("@copies_1", FormatString(lineSegments[lineSegmentCounter + 15].ToString())),
                                               new SqlParameter("@copies_2", FormatString(lineSegments[lineSegmentCounter + 16].ToString())),
                                               new SqlParameter("@copies_3", FormatString(lineSegments[lineSegmentCounter + 17].ToString())),
                                               new SqlParameter("@copies_4", FormatString(lineSegments[lineSegmentCounter + 18].ToString())),
                                               new SqlParameter("@copies_5", FormatString(lineSegments[lineSegmentCounter + 19].ToString())),
                                               new SqlParameter("@copies_6", FormatString(lineSegments[lineSegmentCounter + 20].ToString())),
                                               new SqlParameter("@copies_7", FormatString(lineSegments[lineSegmentCounter + 21].ToString())));

                                    break;
                                case "S2":
                                    S2Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Subscription_S2",
                                               new SqlParameter("@loads_id", loadsId),
                                               new SqlParameter("@pbs_record_number", lineNumber),
                                               new SqlParameter("@segment_instance", SGCount),
                                               new SqlParameter("@original_start_date", FormatDateTime(lineSegments[lineSegmentCounter + 1].ToString())),
                                               new SqlParameter("@source_of_last_start", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                               new SqlParameter("@subsource_of_last_start", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                               new SqlParameter("@credit_status", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                               new SqlParameter("@occupant_type", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                               new SqlParameter("@census_tract", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                               new SqlParameter("@dwelling_type", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                               new SqlParameter("@abc_zone", FormatString(lineSegments[lineSegmentCounter + 8].ToString())),
                                               new SqlParameter("@method_of_renewal_delivery", FormatString(lineSegments[lineSegmentCounter + 1].ToString())));
                                    break;
                                case "EG":
                                    EGCount++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_End_EG",
                                               new SqlParameter("@loads_id", loadsId),
                                               new SqlParameter("@pbs_record_number", lineNumber),
                                               new SqlParameter("@segment_instance", EGCount),
                                               new SqlParameter("@number_of_subscribers_in_group", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                               new SqlParameter("@last_field", (object)DBNull.Value));
                                    break;
                                case "F1":
                                    F1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Additional_Fields_F1",
                                               new SqlParameter("@loads_id", loadsId),
                                               new SqlParameter("@pbs_record_number", lineNumber),
                                               new SqlParameter("@segment_instance", F1Count),
                                               new SqlParameter("@renewal_delivery_override_code", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                               new SqlParameter("@route_walk_sequence", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                               new SqlParameter("@trip_walk_sequence", FormatString(lineSegments[lineSegmentCounter + 3].ToString())));
                                    break;
                                case "B1":
                                    B1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Bill_To_B1",
                                               new SqlParameter("@loads_id", loadsId),
                                               new SqlParameter("@pbs_record_number", lineNumber),
                                               new SqlParameter("@segment_instance", B1Count),
                                               new SqlParameter("@bill_to_name", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                               new SqlParameter("@bill_to_address_1", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                               new SqlParameter("@bill_to_address_2", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                               new SqlParameter("@bill_to_address_3", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                               new SqlParameter("@zip", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                               new SqlParameter("@zip_extension", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                               new SqlParameter("@delivery_point_code", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                               new SqlParameter("@po_route_type", FormatString(lineSegments[lineSegmentCounter + 8].ToString())),
                                               new SqlParameter("@bill_to_indicator", FormatString(lineSegments[lineSegmentCounter + 9].ToString())),
                                               new SqlParameter("@bill_to_occupant_id", FormatString(lineSegments[lineSegmentCounter + 10].ToString())),
                                               new SqlParameter("@bill_to_address_id", FormatString(lineSegments[lineSegmentCounter + 11].ToString())),
                                               new SqlParameter("@bill_to_full_billing_name", FormatString(lineSegments[lineSegmentCounter + 12].ToString())),
                                               new SqlParameter("@bill_to_other_name", FormatString(lineSegments[lineSegmentCounter + 13].ToString())),
                                               new SqlParameter("@imb", FormatString(lineSegments[lineSegmentCounter + 14].ToString())),
                                               new SqlParameter("@encoded_imb_mixed_case", FormatString(lineSegments[lineSegmentCounter + 15].ToString())),
                                               new SqlParameter("@encoded_imb_uppercase", FormatString(lineSegments[lineSegmentCounter + 16].ToString())),
                                               new SqlParameter("@bill_to_address_isonline", FormatString(lineSegments[lineSegmentCounter + 17].ToString())));
                                    break;
                                case "C1":
                                    C1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Carrier_C1",
                                                  new SqlParameter("@loads_id", loadsId),
                                                 new SqlParameter("@pbs_record_number", lineNumber),
                                                 new SqlParameter("@segment_instance", C1Count),
                                                 new SqlParameter("@carrier_name", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                 new SqlParameter("@carrier_home_area_code", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                 new SqlParameter("@carrier_home_phone", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                 new SqlParameter("@district_manager", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                 new SqlParameter("@zone_manager", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                 new SqlParameter("@regional_manager", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                                 new SqlParameter("@area_manager", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                                 new SqlParameter("@depot", FormatString(lineSegments[lineSegmentCounter + 8].ToString())),
                                                 new SqlParameter("@depot_drop_order", FormatString(lineSegments[lineSegmentCounter + 9].ToString())),
                                                 new SqlParameter("@truck", FormatString(lineSegments[lineSegmentCounter + 10].ToString())),
                                                 new SqlParameter("@route_drop_sequence", FormatString(lineSegments[lineSegmentCounter + 11].ToString())));
                                    break;
                                case "D1":
                                    D1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Delivery_D1",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", D1Count),
                                                    new SqlParameter("@deliver_to_name", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@deliver_to_address_1", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@deliver_to_address_2", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@deliver_to_address_3", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@zip", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@zip_extension", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@delivery_point_code", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@newspaper_delivery_route", FormatString(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@route_type", FormatString(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@subscription_home_area_code", FormatString(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@subscription_home_phone", FormatString(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@trip_id", FormatString(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@full_delivery_name", FormatString(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@other_name", FormatString(lineSegments[lineSegmentCounter + 14].ToString())));
                                    break;
                                case "X1":
                                    X1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Demographic_X1",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", X1Count),
                                                    new SqlParameter("@demographic_type", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@demographic_question", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@demographic_answer", FormatString(lineSegments[lineSegmentCounter + 3].ToString())));
                                    break;
                                case "E1":
                                    E1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Expire_E1",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", E1Count),
                                                    new SqlParameter("@transaction_date", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@transaction_description", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@reason_code", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@comments", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@value_of_adjustment", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@days_adjusted", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@transfer_amount", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@transfer_days", FormatString(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@payment_amount", FormatNumber(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@tip_amount", FormatNumber(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@coupon_amount", FormatNumber(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@payment_adjustment_description", FormatString(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@payment_adjustment_amount", FormatNumber(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@update_expire_adjustment_amount", FormatNumber(lineSegments[lineSegmentCounter + 14].ToString())),
                                                    new SqlParameter("@donation_amount", FormatNumber(lineSegments[lineSegmentCounter + 15].ToString())));
                                    break;
                                case "G1":
                                    G1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G1",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", G1Count),
                                                    new SqlParameter("@grace_owed_amount", FormatNumber(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@city_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@county_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@state_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@country_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@total_grace_owed_amount", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())));
                                    break;
                                case "G2":
                                    G2Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G2",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", G2Count),
                                                    new SqlParameter("@tran_type_id", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@tran_date", FormatDateTime(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@renewal_description", FormatString(lineSegments[lineSegmentCounter + 3].ToString())));
                                    break;
                                case "G3":
                                    G3Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G3",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", G3Count),
                                                    new SqlParameter("@grace_owed_amount", FormatNumber(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@city_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@county_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@state_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@country_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@total_grace_owed_amount", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())));
                                    break;
                                case "G4":
                                    G4Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G4",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", G4Count),
                                                    new SqlParameter("@grace_owed_amount", FormatNumber(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@city_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@county_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@state_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@country_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@total_grace_owed_amount", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())));
                                    break;
                                case "Z0":
                                    Z0Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Marketing_Z0",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", Z0Count),
                                                    new SqlParameter("@rate_code", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@rate_code_description", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@marketing_term_length", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@end_date", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@amount", FormatNumber(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@discount_amount", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@city_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@county_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@state_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@country_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@sunday_rate", FormatNumber(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@monday_rate", FormatNumber(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@tuesday_rate", FormatNumber(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@wednesday_rate", FormatNumber(lineSegments[lineSegmentCounter + 14].ToString())),
                                                    new SqlParameter("@thursday_rate", FormatNumber(lineSegments[lineSegmentCounter + 15].ToString())),
                                                    new SqlParameter("@friday_rate", FormatNumber(lineSegments[lineSegmentCounter + 16].ToString())),
                                                    new SqlParameter("@saturday_rate", FormatNumber(lineSegments[lineSegmentCounter + 17].ToString())),
                                                    new SqlParameter("@sunday_discount", FormatNumber(lineSegments[lineSegmentCounter + 18].ToString())),
                                                    new SqlParameter("@monday_discount", FormatNumber(lineSegments[lineSegmentCounter + 19].ToString())),
                                                    new SqlParameter("@tuesday_discount", FormatNumber(lineSegments[lineSegmentCounter + 20].ToString())),
                                                    new SqlParameter("@wednesday_discount", FormatNumber(lineSegments[lineSegmentCounter + 21].ToString())),
                                                    new SqlParameter("@thursday_discount", FormatNumber(lineSegments[lineSegmentCounter + 22].ToString())),
                                                    new SqlParameter("@friday_discount", FormatNumber(lineSegments[lineSegmentCounter + 23].ToString())),
                                                    new SqlParameter("@saturday_discount", FormatNumber(lineSegments[lineSegmentCounter + 24].ToString())));
                                    break;
                                case "Z1":
                                    Z1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Marketing_Z1",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", Z1Count),
                                                    new SqlParameter("@rate_code", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@rate_code_description", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@marketing_term_length", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@end_date", FormatDateTime(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@amount", FormatNumber(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@discount_amount", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@city_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@county_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@state_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@country_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@sunday_rate", FormatNumber(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@monday_rate", FormatNumber(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@tuesday_rate", FormatNumber(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@wednesday_rate", FormatNumber(lineSegments[lineSegmentCounter + 14].ToString())),
                                                    new SqlParameter("@thursday_rate", FormatNumber(lineSegments[lineSegmentCounter + 15].ToString())),
                                                    new SqlParameter("@friday_rate", FormatNumber(lineSegments[lineSegmentCounter + 16].ToString())),
                                                    new SqlParameter("@saturday_rate", FormatNumber(lineSegments[lineSegmentCounter + 17].ToString())),
                                                    new SqlParameter("@sunday_discount", FormatNumber(lineSegments[lineSegmentCounter + 18].ToString())),
                                                    new SqlParameter("@monday_discount", FormatNumber(lineSegments[lineSegmentCounter + 19].ToString())),
                                                    new SqlParameter("@tuesday_discount", FormatNumber(lineSegments[lineSegmentCounter + 20].ToString())),
                                                    new SqlParameter("@wednesday_discount", FormatNumber(lineSegments[lineSegmentCounter + 21].ToString())),
                                                    new SqlParameter("@thursday_discount", FormatNumber(lineSegments[lineSegmentCounter + 22].ToString())),
                                                    new SqlParameter("@friday_discount", FormatNumber(lineSegments[lineSegmentCounter + 23].ToString())),
                                                    new SqlParameter("@saturday_discount", FormatNumber(lineSegments[lineSegmentCounter + 24].ToString())));
                                    break;
                                case "Z2":
                                    Z2Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Marketing_Z2",
                                                   new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", Z2Count),
                                                    new SqlParameter("@rate_code", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@rate_code_description", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@marketing_term_length", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@end_date", FormatDateTime(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@amount", FormatNumber(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@discount_amount", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@city_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@county_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@state_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@country_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@sunday_rate", FormatNumber(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@monday_rate", FormatNumber(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@tuesday_rate", FormatNumber(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@wednesday_rate", FormatNumber(lineSegments[lineSegmentCounter + 14].ToString())),
                                                    new SqlParameter("@thursday_rate", FormatNumber(lineSegments[lineSegmentCounter + 15].ToString())),
                                                    new SqlParameter("@friday_rate", FormatNumber(lineSegments[lineSegmentCounter + 16].ToString())),
                                                    new SqlParameter("@saturday_rate", FormatNumber(lineSegments[lineSegmentCounter + 17].ToString())),
                                                    new SqlParameter("@sunday_discount", FormatNumber(lineSegments[lineSegmentCounter + 18].ToString())),
                                                    new SqlParameter("@monday_discount", FormatNumber(lineSegments[lineSegmentCounter + 19].ToString())),
                                                    new SqlParameter("@tuesday_discount", FormatNumber(lineSegments[lineSegmentCounter + 20].ToString())),
                                                    new SqlParameter("@wednesday_discount", FormatNumber(lineSegments[lineSegmentCounter + 21].ToString())),
                                                    new SqlParameter("@thursday_discount", FormatNumber(lineSegments[lineSegmentCounter + 22].ToString())),
                                                    new SqlParameter("@friday_discount", FormatNumber(lineSegments[lineSegmentCounter + 23].ToString())),
                                                    new SqlParameter("@saturday_discount", FormatNumber(lineSegments[lineSegmentCounter + 24].ToString())));
                                    break;
                                case "M1":
                                    M1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Message_M1",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", M1Count),
                                                    new SqlParameter("@copy_message_1", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@copy_message_2", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@country_tax_message", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@state_tax_message", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@county_tax_message", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@city_tax_message", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@purchase_order_number", FormatString(lineSegments[lineSegmentCounter + 7].ToString())));
                                    break;
                                case "M2":
                                    M2Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Message_M2",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", M2Count),
                                                    new SqlParameter("@renewal_period_message_1", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@renewal_period_message_2", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@general_message", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@rate_code_message_1", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@rate_code_message_2", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@rate_code_message_3", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@rate_code_message_4", FormatString(lineSegments[lineSegmentCounter + 7].ToString())));
                                    break;
                                case "P1":
                                    P1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Payment_P1",
                                                   new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", P1Count),
                                                    new SqlParameter("@payment_amount", FormatNumber(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@payment_length", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@payment_term", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@effective_date", FormatDateTime(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@rate_code", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@payment_transaction_date", FormatDateTime(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@transaction_type", FormatString(lineSegments[lineSegmentCounter + 7].ToString())));
                                    break;
                                case "P2":
                                    P2Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Payment_P2",
                                                   new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", P2Count),
                                                    new SqlParameter("@payment_amount", FormatNumber(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@payment_length", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@payment_term", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@effective_date", FormatDateTime(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@rate_code", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@payment_transaction_date", FormatDateTime(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@transaction_type", FormatString(lineSegments[lineSegmentCounter + 7].ToString())));
                                    break;
                                case "R1":
                                    R1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Rate_R1",
                                                   new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", R1Count),
                                                    new SqlParameter("@rate_code", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@subscription_rate_description_1", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@subscription_rate_before_tax_1", FormatNumber(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@subscription_rate_1", FormatNumber(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@new_expire_date_1", FormatDateTime(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@discount_amount_1", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@rate_by_day_1_1", FormatNumber(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@rate_by_day_1_2", FormatNumber(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@rate_by_day_1_3", FormatNumber(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@rate_by_day_1_4", FormatNumber(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@rate_by_day_1_5", FormatNumber(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@rate_by_day_1_6", FormatNumber(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@rate_by_day_1_7", FormatNumber(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@cost_by_day_1_1", FormatNumber(lineSegments[lineSegmentCounter + 14].ToString())),
                                                    new SqlParameter("@cost_by_day_1_2", FormatNumber(lineSegments[lineSegmentCounter + 15].ToString())),
                                                    new SqlParameter("@cost_by_day_1_3", FormatNumber(lineSegments[lineSegmentCounter + 16].ToString())),
                                                    new SqlParameter("@cost_by_day_1_4", FormatNumber(lineSegments[lineSegmentCounter + 17].ToString())),
                                                    new SqlParameter("@cost_by_day_1_5", FormatNumber(lineSegments[lineSegmentCounter + 18].ToString())),
                                                    new SqlParameter("@cost_by_day_1_6", FormatNumber(lineSegments[lineSegmentCounter + 19].ToString())),
                                                    new SqlParameter("@cost_by_day_1_7", FormatNumber(lineSegments[lineSegmentCounter + 20].ToString())),
                                                    new SqlParameter("@discount_by_day_1_1", FormatNumber(lineSegments[lineSegmentCounter + 21].ToString())),
                                                    new SqlParameter("@discount_by_day_1_2", FormatNumber(lineSegments[lineSegmentCounter + 22].ToString())),
                                                    new SqlParameter("@discount_by_day_1_3", FormatNumber(lineSegments[lineSegmentCounter + 23].ToString())),
                                                    new SqlParameter("@discount_by_day_1_4", FormatNumber(lineSegments[lineSegmentCounter + 24].ToString())),
                                                    new SqlParameter("@discount_by_day_1_5", FormatNumber(lineSegments[lineSegmentCounter + 25].ToString())),
                                                    new SqlParameter("@discount_by_day_1_6", FormatNumber(lineSegments[lineSegmentCounter + 26].ToString())),
                                                    new SqlParameter("@discount_by_day_1_7", FormatNumber(lineSegments[lineSegmentCounter + 27].ToString())),
                                                    new SqlParameter("@subscription_rate_description_2", FormatString(lineSegments[lineSegmentCounter + 28].ToString())),
                                                    new SqlParameter("@subscription_rate_before_tax_2", FormatNumber(lineSegments[lineSegmentCounter + 29].ToString())),
                                                    new SqlParameter("@subscription_rate_2", FormatNumber(lineSegments[lineSegmentCounter + 30].ToString())),
                                                    new SqlParameter("@new_expire_date_2", FormatDateTime(lineSegments[lineSegmentCounter + 31].ToString())),
                                                    new SqlParameter("@discount_amount_2", FormatNumber(lineSegments[lineSegmentCounter + 32].ToString())),
                                                    new SqlParameter("@rate_by_day_2_1", FormatNumber(lineSegments[lineSegmentCounter + 33].ToString())),
                                                    new SqlParameter("@rate_by_day_2_2", FormatNumber(lineSegments[lineSegmentCounter + 34].ToString())),
                                                    new SqlParameter("@rate_by_day_2_3", FormatNumber(lineSegments[lineSegmentCounter + 35].ToString())),
                                                    new SqlParameter("@rate_by_day_2_4", FormatNumber(lineSegments[lineSegmentCounter + 36].ToString())),
                                                    new SqlParameter("@rate_by_day_2_5", FormatNumber(lineSegments[lineSegmentCounter + 37].ToString())),
                                                    new SqlParameter("@rate_by_day_2_6", FormatNumber(lineSegments[lineSegmentCounter + 38].ToString())),
                                                    new SqlParameter("@rate_by_day_2_7", FormatNumber(lineSegments[lineSegmentCounter + 39].ToString())),
                                                    new SqlParameter("@cost_by_day_2_1", FormatNumber(lineSegments[lineSegmentCounter + 40].ToString())),
                                                    new SqlParameter("@cost_by_day_2_2", FormatNumber(lineSegments[lineSegmentCounter + 41].ToString())),
                                                    new SqlParameter("@cost_by_day_2_3", FormatNumber(lineSegments[lineSegmentCounter + 42].ToString())),
                                                    new SqlParameter("@cost_by_day_2_4", FormatNumber(lineSegments[lineSegmentCounter + 43].ToString())),
                                                    new SqlParameter("@cost_by_day_2_5", FormatNumber(lineSegments[lineSegmentCounter + 44].ToString())),
                                                    new SqlParameter("@cost_by_day_2_6", FormatNumber(lineSegments[lineSegmentCounter + 45].ToString())),
                                                    new SqlParameter("@cost_by_day_2_7", FormatNumber(lineSegments[lineSegmentCounter + 46].ToString())),
                                                    new SqlParameter("@discount_by_day_2_1", FormatNumber(lineSegments[lineSegmentCounter + 47].ToString())),
                                                    new SqlParameter("@discount_by_day_2_2", FormatNumber(lineSegments[lineSegmentCounter + 48].ToString())),
                                                    new SqlParameter("@discount_by_day_2_3", FormatNumber(lineSegments[lineSegmentCounter + 49].ToString())),
                                                    new SqlParameter("@discount_by_day_2_4", FormatNumber(lineSegments[lineSegmentCounter + 50].ToString())),
                                                    new SqlParameter("@discount_by_day_2_5", FormatNumber(lineSegments[lineSegmentCounter + 51].ToString())),
                                                    new SqlParameter("@discount_by_day_2_6", FormatNumber(lineSegments[lineSegmentCounter + 52].ToString())),
                                                    new SqlParameter("@discount_by_day_2_7", FormatNumber(lineSegments[lineSegmentCounter + 53].ToString())),
                                                    new SqlParameter("@subscription_rate_description_3", FormatString(lineSegments[lineSegmentCounter + 54].ToString())),
                                                    new SqlParameter("@subscription_rate_before_tax_3", FormatNumber(lineSegments[lineSegmentCounter + 55].ToString())),
                                                    new SqlParameter("@subscription_rate_3", FormatNumber(lineSegments[lineSegmentCounter + 56].ToString())),
                                                    new SqlParameter("@new_expire_date_3", FormatDateTime(lineSegments[lineSegmentCounter + 57].ToString())),
                                                    new SqlParameter("@discount_amount_3", FormatNumber(lineSegments[lineSegmentCounter + 58].ToString())),
                                                    new SqlParameter("@rate_by_day_3_1", FormatNumber(lineSegments[lineSegmentCounter + 59].ToString())),
                                                    new SqlParameter("@rate_by_day_3_2", FormatNumber(lineSegments[lineSegmentCounter + 60].ToString())),
                                                    new SqlParameter("@rate_by_day_3_3", FormatNumber(lineSegments[lineSegmentCounter + 61].ToString())),
                                                    new SqlParameter("@rate_by_day_3_4", FormatNumber(lineSegments[lineSegmentCounter + 62].ToString())),
                                                    new SqlParameter("@rate_by_day_3_5", FormatNumber(lineSegments[lineSegmentCounter + 63].ToString())),
                                                    new SqlParameter("@rate_by_day_3_6", FormatNumber(lineSegments[lineSegmentCounter + 64].ToString())),
                                                    new SqlParameter("@rate_by_day_3_7", FormatNumber(lineSegments[lineSegmentCounter + 65].ToString())),
                                                    new SqlParameter("@cost_by_day_3_1", FormatNumber(lineSegments[lineSegmentCounter + 66].ToString())),
                                                    new SqlParameter("@cost_by_day_3_2", FormatNumber(lineSegments[lineSegmentCounter + 67].ToString())),
                                                    new SqlParameter("@cost_by_day_3_3", FormatNumber(lineSegments[lineSegmentCounter + 68].ToString())),
                                                    new SqlParameter("@cost_by_day_3_4", FormatNumber(lineSegments[lineSegmentCounter + 69].ToString())),
                                                    new SqlParameter("@cost_by_day_3_5", FormatNumber(lineSegments[lineSegmentCounter + 70].ToString())),
                                                    new SqlParameter("@cost_by_day_3_6", FormatNumber(lineSegments[lineSegmentCounter + 71].ToString())),
                                                    new SqlParameter("@cost_by_day_3_7", FormatNumber(lineSegments[lineSegmentCounter + 72].ToString())),
                                                    new SqlParameter("@discount_by_day_3_1", FormatNumber(lineSegments[lineSegmentCounter + 73].ToString())),
                                                    new SqlParameter("@discount_by_day_3_2", FormatNumber(lineSegments[lineSegmentCounter + 74].ToString())),
                                                    new SqlParameter("@discount_by_day_3_3", FormatNumber(lineSegments[lineSegmentCounter + 75].ToString())),
                                                    new SqlParameter("@discount_by_day_3_4", FormatNumber(lineSegments[lineSegmentCounter + 76].ToString())),
                                                    new SqlParameter("@discount_by_day_3_5", FormatNumber(lineSegments[lineSegmentCounter + 77].ToString())),
                                                    new SqlParameter("@discount_by_day_3_6", FormatNumber(lineSegments[lineSegmentCounter + 78].ToString())),
                                                    new SqlParameter("@discount_by_day_3_7", FormatNumber(lineSegments[lineSegmentCounter + 79].ToString())),
                                                    new SqlParameter("@subscription_rate_description_4", FormatString(lineSegments[lineSegmentCounter + 80].ToString())),
                                                    new SqlParameter("@subscription_rate_before_tax_4", FormatNumber(lineSegments[lineSegmentCounter + 81].ToString())),
                                                    new SqlParameter("@subscription_rate_4", FormatNumber(lineSegments[lineSegmentCounter + 82].ToString())),
                                                    new SqlParameter("@new_expire_date_4", FormatDateTime(lineSegments[lineSegmentCounter + 83].ToString())),
                                                    new SqlParameter("@discount_amount_4", FormatNumber(lineSegments[lineSegmentCounter + 84].ToString())),
                                                    new SqlParameter("@rate_by_day_4_1", FormatNumber(lineSegments[lineSegmentCounter + 85].ToString())),
                                                    new SqlParameter("@rate_by_day_4_2", FormatNumber(lineSegments[lineSegmentCounter + 86].ToString())),
                                                    new SqlParameter("@rate_by_day_4_3", FormatNumber(lineSegments[lineSegmentCounter + 87].ToString())),
                                                    new SqlParameter("@rate_by_day_4_4", FormatNumber(lineSegments[lineSegmentCounter + 88].ToString())),
                                                    new SqlParameter("@rate_by_day_4_5", FormatNumber(lineSegments[lineSegmentCounter + 89].ToString())),
                                                    new SqlParameter("@rate_by_day_4_6", FormatNumber(lineSegments[lineSegmentCounter + 90].ToString())),
                                                    new SqlParameter("@rate_by_day_4_7", FormatNumber(lineSegments[lineSegmentCounter + 91].ToString())),
                                                    new SqlParameter("@cost_by_day_4_1", FormatNumber(lineSegments[lineSegmentCounter + 92].ToString())),
                                                    new SqlParameter("@cost_by_day_4_2", FormatNumber(lineSegments[lineSegmentCounter + 93].ToString())),
                                                    new SqlParameter("@cost_by_day_4_3", FormatNumber(lineSegments[lineSegmentCounter + 94].ToString())),
                                                    new SqlParameter("@cost_by_day_4_4", FormatNumber(lineSegments[lineSegmentCounter + 95].ToString())),
                                                    new SqlParameter("@cost_by_day_4_5", FormatNumber(lineSegments[lineSegmentCounter + 96].ToString())),
                                                    new SqlParameter("@cost_by_day_4_6", FormatNumber(lineSegments[lineSegmentCounter + 97].ToString())),
                                                    new SqlParameter("@cost_by_day_4_7", FormatNumber(lineSegments[lineSegmentCounter + 98].ToString())),
                                                    new SqlParameter("@discount_by_day_4_1", FormatNumber(lineSegments[lineSegmentCounter + 99].ToString())),
                                                    new SqlParameter("@discount_by_day_4_2", FormatNumber(lineSegments[lineSegmentCounter + 100].ToString())),
                                                    new SqlParameter("@discount_by_day_4_3", FormatNumber(lineSegments[lineSegmentCounter + 101].ToString())),
                                                    new SqlParameter("@discount_by_day_4_4", FormatNumber(lineSegments[lineSegmentCounter + 102].ToString())),
                                                    new SqlParameter("@discount_by_day_4_5", FormatNumber(lineSegments[lineSegmentCounter + 103].ToString())),
                                                    new SqlParameter("@discount_by_day_4_6", FormatNumber(lineSegments[lineSegmentCounter + 104].ToString())),
                                                    new SqlParameter("@discount_by_day_4_7", FormatNumber(lineSegments[lineSegmentCounter + 105].ToString())),
                                                    new SqlParameter("@grace_owed_amount_not_including_taxes", FormatNumber(lineSegments[lineSegmentCounter + 106].ToString())),
                                                    new SqlParameter("@grace_owed_amount_including_taxes", FormatNumber(lineSegments[lineSegmentCounter + 107].ToString())),
                                                    new SqlParameter("@grace_owed_city_tax", FormatNumber(lineSegments[lineSegmentCounter + 108].ToString())),
                                                    new SqlParameter("@grace_owed_county_tax", FormatNumber(lineSegments[lineSegmentCounter + 109].ToString())),
                                                    new SqlParameter("@grace_owed_state_tax", FormatNumber(lineSegments[lineSegmentCounter + 110].ToString())),
                                                    new SqlParameter("@grace_owed_country_tax", FormatNumber(lineSegments[lineSegmentCounter + 111].ToString())),
                                                    new SqlParameter("@subscribers_unallocated_balance", FormatNumber(lineSegments[lineSegmentCounter + 112].ToString())));
                                    break;
                                case "R2":
                                    R2Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Rate_R2",
                                                   new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", R2Count),
                                                    new SqlParameter("@combo_id", FormatString(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@subscription_number", FormatString(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@product_id", FormatString(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@delivery_method", FormatString(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@delivery_schedule", FormatString(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@rate_code", FormatString(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@subscription_rate_description", FormatString(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@subscription_rate_before_tax", FormatNumber(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@subscription_rate", FormatNumber(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@total_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@city_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@county_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@state_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@country_tax_amount", FormatNumber(lineSegments[lineSegmentCounter + 14].ToString())),
                                                    new SqlParameter("@delivery_fee_amount", FormatNumber(lineSegments[lineSegmentCounter + 15].ToString())),
                                                    new SqlParameter("@new_expire_date", FormatDateTime(lineSegments[lineSegmentCounter + 16].ToString())));
                                    break;
                                case "T1":
                                    T1Count++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Tax_T1",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", T1Count),
                                                    new SqlParameter("@country_tax_amount_1", FormatNumber(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@state_tax_amount_1", FormatNumber(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@county_tax_amount_1", FormatNumber(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@city_tax_amount_1", FormatNumber(lineSegments[lineSegmentCounter + 4].ToString())),
                                                    new SqlParameter("@country_tax_amount_2", FormatNumber(lineSegments[lineSegmentCounter + 5].ToString())),
                                                    new SqlParameter("@state_tax_amount_2", FormatNumber(lineSegments[lineSegmentCounter + 6].ToString())),
                                                    new SqlParameter("@county_tax_amount_2", FormatNumber(lineSegments[lineSegmentCounter + 7].ToString())),
                                                    new SqlParameter("@city_tax_amount_2", FormatNumber(lineSegments[lineSegmentCounter + 8].ToString())),
                                                    new SqlParameter("@country_tax_amount_3", FormatNumber(lineSegments[lineSegmentCounter + 9].ToString())),
                                                    new SqlParameter("@state_tax_amount_3", FormatNumber(lineSegments[lineSegmentCounter + 10].ToString())),
                                                    new SqlParameter("@county_tax_amount_3", FormatNumber(lineSegments[lineSegmentCounter + 11].ToString())),
                                                    new SqlParameter("@city_tax_amount_3", FormatNumber(lineSegments[lineSegmentCounter + 12].ToString())),
                                                    new SqlParameter("@country_tax_amount_4", FormatNumber(lineSegments[lineSegmentCounter + 13].ToString())),
                                                    new SqlParameter("@state_tax_amount_4", FormatNumber(lineSegments[lineSegmentCounter + 14].ToString())),
                                                    new SqlParameter("@county_tax_amount_4", FormatNumber(lineSegments[lineSegmentCounter + 15].ToString())),
                                                    new SqlParameter("@city_tax_amount_4", FormatNumber(lineSegments[lineSegmentCounter + 16].ToString())));
                                    break;
                                case "TC":
                                    TCCount++;

                                    ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Transportation_TC",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@pbs_record_number", lineNumber),
                                                    new SqlParameter("@segment_instance", TCCount),
                                                    new SqlParameter("@transportation_cost_1", FormatNumber(lineSegments[lineSegmentCounter + 1].ToString())),
                                                    new SqlParameter("@transportation_cost_2", FormatNumber(lineSegments[lineSegmentCounter + 2].ToString())),
                                                    new SqlParameter("@transportation_cost_3", FormatNumber(lineSegments[lineSegmentCounter + 3].ToString())),
                                                    new SqlParameter("@transportation_cost_4", FormatNumber(lineSegments[lineSegmentCounter + 4].ToString())));
                                    break;
                            }

                        }

                        lineSegmentCounter++;
                    }

                    processedSegmentCount += (F1Count + B1Count + C1Count + D1Count + X1Count + EGCount + E1Count + G1Count + G2Count + G3Count + G4Count + Z0Count + Z1Count + Z2Count + M1Count + M2Count + P1Count + P2Count + R1Count + R2Count + SGCount + S1Count + S2Count + T1Count + TCCount);

                }

            }

            WriteToJobLog(JobLogMessageType.INFO, $"{processedSegmentCount} total segments read.");

            ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Update_Loads_Count",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pintLoadCount", processedSegmentCount),
                                        new SqlParameter("@pflgSuccessfulLoad", 1));

        }

        public override void SetupJob()
        {
            JobName = "Office Pay";
            JobDescription = @"Parses an X12-like file dumping the records into the local database. Each segment type matches the table suffix.";
            AppConfigSectionName = "OfficePay";
        }
    }
}
