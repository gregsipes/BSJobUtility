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

                //  List<string> processedFiles = new List<string>();

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

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Update_Loads_Backup",
                                        new SqlParameter("@pintLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            //Int32 routeDetailCounter = 0;
            //Int32 advanceDetailCounter = 0;
            //Int32 advanceTotalCounter = 0;
            //Int32 TMDetailCounter = 0;
            //Int32 TMTotalCounter = 0;
            //Int32 truckTotalCounter = 0;
            //Int32 ignoredCounter = 0;

            Int32 lineNumber = 0;
    

            //DateTime? runDate = null;
            //String runType = "";

            foreach (string line in fileContents)
            {
                lineNumber++;

                if (line != null && line.Trim().Length > 0)
                {
                    List<string> lineSegments = line.Split('|').ToList();
                    List<string> segmentNames = new List<string>() { "F1", "B1", "C1", "D1", "X1", "EG", "E1", "G1", "G2", "G3",
                                                                     "G4", "Z0", "Z1", "Z2", "M1", "M2", "P1", "P2", "R1", "R2",
                                                                     "SG", "S1", "S2", "T1", "TC" };
                    string currentSegmentName = lineSegments[0];
                    Int32 lineSegmentCounter = 0;

                    foreach (string lineSegment in lineSegments)
                    {
                        lineSegmentCounter++;

                        //check to see if we are starting a new segment
                        if (segmentNames.Where(s => s == lineSegment).FirstOrDefault() != null)
                            currentSegmentName = lineSegment;

                        switch (currentSegmentName)
                        {
                            case "SG":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Start_SG",
                                               new SqlParameter("@loads_id", loadsId),
                                               new SqlParameter("@pbs_record_number", lineNumber),
                                               new SqlParameter("@segment_instance", 1),
                                               new SqlParameter("@bill_to_name", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                               new SqlParameter("@bill_to_address_1", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                               new SqlParameter("@bill_to_address_2", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                               new SqlParameter("@bill_to_address_3", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                               new SqlParameter("@zip", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                               new SqlParameter("@zip_extension", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                               new SqlParameter("@delivery_point_code", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                               new SqlParameter("@route", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                               new SqlParameter("@imb", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                               new SqlParameter("@encoded_imb_mixed_case", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                               new SqlParameter("@encoded_imb_uppercase", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                               new SqlParameter("@last_field", (object)DBNull.Value));
                                break;
                            case "S1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Subscription_S1",
                                           new SqlParameter("@loads_id", loadsId),
                                           new SqlParameter("@pbs_record_number", lineNumber),
                                           new SqlParameter("@segment_instance", 1),
                                           new SqlParameter("@subscription_number", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                           new SqlParameter("@delivery_schedule", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                           new SqlParameter("@expire_date", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                           new SqlParameter("@end_grace_date", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                           new SqlParameter("@days_of_week_1", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                           new SqlParameter("@days_of_week_2", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                           new SqlParameter("@publication_name", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                           new SqlParameter("@lockbox_scanline_data", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                           new SqlParameter("@renewal_run_date", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                           new SqlParameter("@renewal_invoice_or_grace", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                           new SqlParameter("@last_renewal_date", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                           new SqlParameter("@renewal_number", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                           new SqlParameter("@start_reason_code", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                           new SqlParameter("@number_of_payments_since_start", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()),
                                           new SqlParameter("@copies_1", lineSegments[lineSegmentCounter + 15].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 15].ToString()),
                                           new SqlParameter("@copies_2", lineSegments[lineSegmentCounter + 16].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 16].ToString()),
                                           new SqlParameter("@copies_3", lineSegments[lineSegmentCounter + 17].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 17].ToString()),
                                           new SqlParameter("@copies_4", lineSegments[lineSegmentCounter + 18].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 18].ToString()),
                                           new SqlParameter("@copies_5", lineSegments[lineSegmentCounter + 19].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 19].ToString()),
                                           new SqlParameter("@copies_6", lineSegments[lineSegmentCounter + 20].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 20].ToString()),
                                           new SqlParameter("@copies_7", lineSegments[lineSegmentCounter + 21].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 21].ToString()));
                                break;
                            case "EG":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_End_EG",
                                           new SqlParameter("@loads_id", loadsId),
                                           new SqlParameter("@pbs_record_number", lineNumber),
                                           new SqlParameter("@segment_instace", 1),
                                           new SqlParameter("@number_of-subscribers_in_group", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                           new SqlParameter("@last_field", (object)DBNull.Value));
                                break;
                            case "F1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Additional_Fields_F1",
                                           new SqlParameter("@loads_id", loadsId),
                                           new SqlParameter("@pbs_record_number", lineNumber),
                                           new SqlParameter("@segment_instance", 1),
                                           new SqlParameter("@renewal_delivery_override_code", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                           new SqlParameter("@route_walk_sequence", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                           new SqlParameter("@trip_walk_sequence", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()));
                                break;
                            case "B1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Bill_To_B1",
                                           new SqlParameter("@loads_id", loadsId),
                                           new SqlParameter("@pbs_record_number", lineNumber),
                                           new SqlParameter("@bill_to_name", 1),
                                           new SqlParameter("@bill_to_address_1", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                           new SqlParameter("@bill_to_address_2", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                           new SqlParameter("@bill_to_address_3", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                           new SqlParameter("@zip", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                           new SqlParameter("@zip_extension", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                           new SqlParameter("@delivery_point_code", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                           new SqlParameter("@pro_route_type", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                           new SqlParameter("@bill_to_indicator", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                           new SqlParameter("@bill_to_occupant_id", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                           new SqlParameter("@bill_to_address_id", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                           new SqlParameter("@bill_to_full_billing_name", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                           new SqlParameter("@bill_to_other_name", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                           new SqlParameter("@imb", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                           new SqlParameter("@encoded_imb_mixed_case", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()),
                                           new SqlParameter("@encoded_imb_uppercase", lineSegments[lineSegmentCounter + 15].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 15].ToString()),
                                           new SqlParameter("@bill_to_address_isonline", lineSegments[lineSegmentCounter + 16].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 16].ToString()));
                                break;
                            case "C1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Carrier_C1",
                                              new SqlParameter("@loads_id", loadsId),
                                             new SqlParameter("@pbs_record_number", lineNumber),
                                             new SqlParameter("@segment_instance", 1),
                                             new SqlParameter("@carrier_name", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                             new SqlParameter("@carrier_home_area_code", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                             new SqlParameter("@carrier_home_phone", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                             new SqlParameter("@district_manager", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                             new SqlParameter("@zone_manager", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                             new SqlParameter("@regional_manager", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                             new SqlParameter("@area_manager", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                             new SqlParameter("@depot", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                             new SqlParameter("@depot_drop_order", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                             new SqlParameter("@truck", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                             new SqlParameter("@route_drop_sequence", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()));
                                break;
                            case "D1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Delivery_D1",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@deliver_to_name", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@deliver_to_address_1", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@deliver_to_address_2", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@deliver_to_address_3", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@zip", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@zip_extension", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@delivery_point_code", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                                new SqlParameter("@newspaper_delivery_route", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                                new SqlParameter("@route_type", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                                new SqlParameter("@subscription_home_area_code", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                                new SqlParameter("@subscription_home_phone", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                                new SqlParameter("@trip_id", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                                new SqlParameter("@full_delivery_name", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                                new SqlParameter("@other_name", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()));
                                break;
                            case "X1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Demographic_X1",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@demographic_type", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@demographic_question", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@demographic_answer", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()));
                                break;
                            case "E1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Expire_E1",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@transaction_date", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@transaction_description", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@reason_code", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@comments", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@value_of_adjustment", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@days_adjusted", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@transfer_amount", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                                new SqlParameter("@transfer_days", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                                new SqlParameter("@payment_amount", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                                new SqlParameter("@tip_amount", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                                new SqlParameter("@coupon_amount", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                                new SqlParameter("@payment_adjustment_description", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                                new SqlParameter("@payment_adjustment_amount", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                                new SqlParameter("@update_expire_adjustment_amount", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()),
                                                new SqlParameter("@donation_amount", lineSegments[lineSegmentCounter + 15].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 15].ToString()));
                                break;
                            case "G1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G1",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@grace_owed_amount", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@city_tax_amount", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@county_tax_amount", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@state_tax_amount", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@country_tax_amount", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@total_grace_owed_amount", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()));
                                break;
                            case "G2":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G2",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@tran_type_id", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@tran_date", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@renewal_description", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@state_tax_amount", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()));
                                break;
                            case "G3":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G3",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@grace_owed_amount", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@city_tax_amount", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@county_tax_amount", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@state_tax_amount", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@country_tax_amount", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@total_grace_owed_amount", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()));
                                break;
                            case "G4":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Grace_G4",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@grace_owed_amount", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@city_tax_amount", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@county_tax_amount", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@state_tax_amount", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@country_tax_amount", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@total_grace_owed_amount", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()));
                                break;
                            case "Z0":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Marketing_Z0",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@rate_code", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@rate_code_description", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@marketing_term_length", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@end_date", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@amount", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@discount_amount", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@city_tax_amount", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                                new SqlParameter("@county_tax_amount", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                                new SqlParameter("@state_tax_amount", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                                new SqlParameter("@country_tax_amount", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                                new SqlParameter("@sunday_rate", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                                new SqlParameter("@monday_rate", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                                new SqlParameter("@tuesday_rate", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                                new SqlParameter("@wednesday_rate", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()),
                                                new SqlParameter("@thursday_rate", lineSegments[lineSegmentCounter + 15].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 15].ToString()),
                                                new SqlParameter("@friday_rate", lineSegments[lineSegmentCounter + 16].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 16].ToString()),
                                                new SqlParameter("@saturday_rate", lineSegments[lineSegmentCounter + 17].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 17].ToString()),
                                                new SqlParameter("@sunday_discount", lineSegments[lineSegmentCounter + 18].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 18].ToString()),
                                                new SqlParameter("@monday_discount", lineSegments[lineSegmentCounter + 19].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 19].ToString()),
                                                new SqlParameter("@tuesday_discount", lineSegments[lineSegmentCounter + 20].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 20].ToString()),
                                                new SqlParameter("@wednesday_discount", lineSegments[lineSegmentCounter + 21].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 21].ToString()),
                                                new SqlParameter("@thursday_discount", lineSegments[lineSegmentCounter + 22].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 22].ToString()),
                                                new SqlParameter("@friday_discount", lineSegments[lineSegmentCounter + 23].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 23].ToString()),
                                                new SqlParameter("@saturday_discount", lineSegments[lineSegmentCounter + 24].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 24].ToString()));
                                break;
                            case "Z1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Marketing_Z1",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@rate_code", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@rate_code_description", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@marketing_term_length", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@end_date", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@amount", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@discount_amount", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@city_tax_amount", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                                new SqlParameter("@county_tax_amount", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                                new SqlParameter("@state_tax_amount", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                                new SqlParameter("@country_tax_amount", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                                new SqlParameter("@sunday_rate", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                                new SqlParameter("@monday_rate", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                                new SqlParameter("@tuesday_rate", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                                new SqlParameter("@wednesday_rate", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()),
                                                new SqlParameter("@thursday_rate", lineSegments[lineSegmentCounter + 15].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 15].ToString()),
                                                new SqlParameter("@friday_rate", lineSegments[lineSegmentCounter + 16].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 16].ToString()),
                                                new SqlParameter("@saturday_rate", lineSegments[lineSegmentCounter + 17].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 17].ToString()),
                                                new SqlParameter("@sunday_discount", lineSegments[lineSegmentCounter + 18].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 18].ToString()),
                                                new SqlParameter("@monday_discount", lineSegments[lineSegmentCounter + 19].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 19].ToString()),
                                                new SqlParameter("@tuesday_discount", lineSegments[lineSegmentCounter + 20].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 20].ToString()),
                                                new SqlParameter("@wednesday_discount", lineSegments[lineSegmentCounter + 21].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 21].ToString()),
                                                new SqlParameter("@thursday_discount", lineSegments[lineSegmentCounter + 22].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 22].ToString()),
                                                new SqlParameter("@friday_discount", lineSegments[lineSegmentCounter + 23].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 23].ToString()),
                                                new SqlParameter("@saturday_discount", lineSegments[lineSegmentCounter + 24].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 24].ToString()));
                                break;
                            case "Z2":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Marketing_Z2",
                                               new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@rate_code", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@rate_code_description", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@marketing_term_length", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@end_date", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@amount", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@discount_amount", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@city_tax_amount", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                                new SqlParameter("@county_tax_amount", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                                new SqlParameter("@state_tax_amount", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                                new SqlParameter("@country_tax_amount", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                                new SqlParameter("@sunday_rate", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                                new SqlParameter("@monday_rate", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                                new SqlParameter("@tuesday_rate", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                                new SqlParameter("@wednesday_rate", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()),
                                                new SqlParameter("@thursday_rate", lineSegments[lineSegmentCounter + 15].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 15].ToString()),
                                                new SqlParameter("@friday_rate", lineSegments[lineSegmentCounter + 16].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 16].ToString()),
                                                new SqlParameter("@saturday_rate", lineSegments[lineSegmentCounter + 17].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 17].ToString()),
                                                new SqlParameter("@sunday_discount", lineSegments[lineSegmentCounter + 18].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 18].ToString()),
                                                new SqlParameter("@monday_discount", lineSegments[lineSegmentCounter + 19].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 19].ToString()),
                                                new SqlParameter("@tuesday_discount", lineSegments[lineSegmentCounter + 20].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 20].ToString()),
                                                new SqlParameter("@wednesday_discount", lineSegments[lineSegmentCounter + 21].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 21].ToString()),
                                                new SqlParameter("@thursday_discount", lineSegments[lineSegmentCounter + 22].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 22].ToString()),
                                                new SqlParameter("@friday_discount", lineSegments[lineSegmentCounter + 23].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 23].ToString()),
                                                new SqlParameter("@saturday_discount", lineSegments[lineSegmentCounter + 24].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 24].ToString()));
                                break;
                            case "M1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Message_M1",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@copy_message_1", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@copy_message_2", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@country_tax_message", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@state_tax_message", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@county_tax_message", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@city_tax_message", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@purchase_order_number", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()));
                                break;
                            case "M2":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Message_M2",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@renewal_period_message_1", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@renewal_period_message_2", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@general_message", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@rate_code_message_1", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@rate_code_message_2", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@rate_code_message_3", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@rate_code_message_4", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()));
                                break;
                            case "P1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Payment_P1",
                                               new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@payment_amount", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@payment_length", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@payment_term", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@effective_date", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@rate_code", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@payment_transaction_date", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@transaction_type", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()));
                                      break;
                            case "P2":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Payment_P2",
                                               new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@payment_amount", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@payment_length", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@payment_term", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@effective_date", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@rate_code", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@payment_transaction_date", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@transaction_type", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()));
                                break;
                            case "R1":
                                ExecuteNonQuery(DatabaseConnectionStringNames.OfficePay, "dbo.Proc_Insert_Rate_R1",
                                               new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@pbs_record_number", lineNumber),
                                                new SqlParameter("@segment_instance", 1),
                                                new SqlParameter("@rate_code", lineSegments[lineSegmentCounter + 1].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 1].ToString()),
                                                new SqlParameter("@subscription_rate_description_1", lineSegments[lineSegmentCounter + 2].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 2].ToString()),
                                                new SqlParameter("@subscription_rate_before_tax_1", lineSegments[lineSegmentCounter + 3].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 3].ToString()),
                                                new SqlParameter("@subscription_rate_1", lineSegments[lineSegmentCounter + 4].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 4].ToString()),
                                                new SqlParameter("@new_expire_date_1", lineSegments[lineSegmentCounter + 5].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 5].ToString()),
                                                new SqlParameter("@discount_amount_1", lineSegments[lineSegmentCounter + 6].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 6].ToString()),
                                                new SqlParameter("@rate_by_day_1_1", lineSegments[lineSegmentCounter + 7].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 7].ToString()),
                                                new SqlParameter("@rate_by_day_1_2", lineSegments[lineSegmentCounter + 8].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 8].ToString()),
                                                new SqlParameter("@rate_by_day_1_3", lineSegments[lineSegmentCounter + 9].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 9].ToString()),
                                                new SqlParameter("@rate_by_day_1_4", lineSegments[lineSegmentCounter + 10].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 10].ToString()),
                                                new SqlParameter("@rate_by_day_1_5", lineSegments[lineSegmentCounter + 11].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 11].ToString()),
                                                new SqlParameter("@rate_by_day_1_6", lineSegments[lineSegmentCounter + 12].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 12].ToString()),
                                                new SqlParameter("@rate_by_day_1_7", lineSegments[lineSegmentCounter + 13].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 13].ToString()),
                                                new SqlParameter("@cost_by_day_1_1", lineSegments[lineSegmentCounter + 14].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 14].ToString()),
                                                new SqlParameter("@cost_by_day_1_2", lineSegments[lineSegmentCounter + 15].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 15].ToString()),
                                                new SqlParameter("@cost_by_day_1_3", lineSegments[lineSegmentCounter + 16].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 16].ToString()),
                                                new SqlParameter("@cost_by_day_1_4", lineSegments[lineSegmentCounter + 17].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 17].ToString()),
                                                new SqlParameter("@cost_by_day_1_5", lineSegments[lineSegmentCounter + 18].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 18].ToString()),
                                                new SqlParameter("@cost_by_day_1_6", lineSegments[lineSegmentCounter + 19].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 19].ToString()),
                                                new SqlParameter("@cost_by_day_1_7", lineSegments[lineSegmentCounter + 20].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 20].ToString()),
                                                new SqlParameter("@discount_by_day_1_1", lineSegments[lineSegmentCounter + 21].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 21].ToString()),
                                                new SqlParameter("@discount_by_day_1_2", lineSegments[lineSegmentCounter + 22].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 22].ToString()),
                                                new SqlParameter("@discount_by_day_1_3", lineSegments[lineSegmentCounter + 23].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 23].ToString()),
                                                new SqlParameter("@discount_by_day_1_4", lineSegments[lineSegmentCounter + 24].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 24].ToString()),
                                                new SqlParameter("@discount_by_day_1_5", lineSegments[lineSegmentCounter + 25].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 25].ToString()),
                                                new SqlParameter("@discount_by_day_1_6", lineSegments[lineSegmentCounter + 26].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 26].ToString()),
                                                new SqlParameter("@discount_by_day_1_7", lineSegments[lineSegmentCounter + 27].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 27].ToString()),
                                                new SqlParameter("@subscription_rate_description_2", lineSegments[lineSegmentCounter + 28].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 28].ToString()),
                                                new SqlParameter("@subscription_rate_before_tax_2", lineSegments[lineSegmentCounter + 29].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 29].ToString()),
                                                new SqlParameter("@subscription_rate_2", lineSegments[lineSegmentCounter + 30].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 30].ToString()),
                                                new SqlParameter("@new_expire_date_2", lineSegments[lineSegmentCounter + 31].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 31].ToString()),
                                                new SqlParameter("@discount_amount_2", lineSegments[lineSegmentCounter + 32].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 32].ToString()),
                                                new SqlParameter("@rate_by_day_2_1", lineSegments[lineSegmentCounter + 33].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 33].ToString()),
                                                new SqlParameter("@rate_by_day_2_2", lineSegments[lineSegmentCounter + 34].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 34].ToString()),
                                                new SqlParameter("@rate_by_day_2_3", lineSegments[lineSegmentCounter + 35].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 35].ToString()),
                                                new SqlParameter("@rate_by_day_2_4", lineSegments[lineSegmentCounter + 36].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 36].ToString()),
                                                new SqlParameter("@rate_by_day_2_5", lineSegments[lineSegmentCounter + 37].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 37].ToString()),
                                                new SqlParameter("@rate_by_day_2_6", lineSegments[lineSegmentCounter + 38].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 38].ToString()),
                                                new SqlParameter("@rate_by_day_2_7", lineSegments[lineSegmentCounter + 39].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 39].ToString()),
                                                new SqlParameter("@cost_by_day_2_1", lineSegments[lineSegmentCounter + 40].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 40].ToString()),
                                                new SqlParameter("@cost_by_day_2_2", lineSegments[lineSegmentCounter + 41].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 41].ToString()),
                                                new SqlParameter("@cost_by_day_2_3", lineSegments[lineSegmentCounter + 42].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 42].ToString()),
                                                new SqlParameter("@cost_by_day_2_4", lineSegments[lineSegmentCounter + 43].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 43].ToString()),
                                                new SqlParameter("@cost_by_day_2_5", lineSegments[lineSegmentCounter + 44].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 44].ToString()),
                                                new SqlParameter("@cost_by_day_2_6", lineSegments[lineSegmentCounter + 45].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 45].ToString()),
                                                new SqlParameter("@cost_by_day_2_7", lineSegments[lineSegmentCounter + 46].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 46].ToString()),
                                                new SqlParameter("@discount_by_day_2_1", lineSegments[lineSegmentCounter + 47].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 47].ToString()),
                                                new SqlParameter("@discount_by_day_2_2", lineSegments[lineSegmentCounter + 48].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 48].ToString()),
                                                new SqlParameter("@discount_by_day_2_3", lineSegments[lineSegmentCounter + 49].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 49].ToString()),
                                                new SqlParameter("@discount_by_day_2_4", lineSegments[lineSegmentCounter + 50].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 50].ToString()),
                                                new SqlParameter("@discount_by_day_2_5", lineSegments[lineSegmentCounter + 51].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 51].ToString()),
                                                new SqlParameter("@discount_by_day_2_6", lineSegments[lineSegmentCounter + 52].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 52].ToString()),
                                                new SqlParameter("@discount_by_day_2_7", lineSegments[lineSegmentCounter + 53].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 53].ToString()),
                                                new SqlParameter("@subscription_rate_description_3", lineSegments[lineSegmentCounter + 54].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 54].ToString()),
                                                new SqlParameter("@subscription_rate_before_tax_3", lineSegments[lineSegmentCounter + 55].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 55].ToString()),
                                                new SqlParameter("@subscription_rate_3", lineSegments[lineSegmentCounter + 56].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 56].ToString()),
                                                new SqlParameter("@new_expire_date_3", lineSegments[lineSegmentCounter + 57].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 57].ToString()),
                                                new SqlParameter("@discount_amount_3", lineSegments[lineSegmentCounter + 58].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 58].ToString()),
                                                new SqlParameter("@rate_by_day_3_1", lineSegments[lineSegmentCounter + 59].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 59].ToString()),
                                                new SqlParameter("@rate_by_day_3_2", lineSegments[lineSegmentCounter + 60].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 60].ToString()),
                                                new SqlParameter("@rate_by_day_3_3", lineSegments[lineSegmentCounter + 61].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 61].ToString()),
                                                new SqlParameter("@rate_by_day_3_4", lineSegments[lineSegmentCounter + 62].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 62].ToString()),
                                                new SqlParameter("@rate_by_day_3_5", lineSegments[lineSegmentCounter + 63].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 63].ToString()),
                                                new SqlParameter("@rate_by_day_3_6", lineSegments[lineSegmentCounter + 64].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 64].ToString()),
                                                new SqlParameter("@rate_by_day_3_7", lineSegments[lineSegmentCounter + 65].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 65].ToString()),
                                                new SqlParameter("@cost_by_day_3_1", lineSegments[lineSegmentCounter + 66].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 66].ToString()),
                                                new SqlParameter("@cost_by_day_3_2", lineSegments[lineSegmentCounter + 67].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 67].ToString()),
                                                new SqlParameter("@cost_by_day_3_3", lineSegments[lineSegmentCounter + 68].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 68].ToString()),
                                                new SqlParameter("@cost_by_day_3_4", lineSegments[lineSegmentCounter + 69].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 69].ToString()),
                                                new SqlParameter("@cost_by_day_3_5", lineSegments[lineSegmentCounter + 70].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 70].ToString()),
                                                new SqlParameter("@cost_by_day_3_6", lineSegments[lineSegmentCounter + 71].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 71].ToString()),
                                                new SqlParameter("@cost_by_day_3_7", lineSegments[lineSegmentCounter + 72].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 72].ToString()),
                                                new SqlParameter("@discount_by_day_3_1", lineSegments[lineSegmentCounter + 73].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 73].ToString()),
                                                new SqlParameter("@discount_by_day_3_2", lineSegments[lineSegmentCounter + 74].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 74].ToString()),
                                                new SqlParameter("@discount_by_day_3_3", lineSegments[lineSegmentCounter + 75].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 75].ToString()),
                                                new SqlParameter("@discount_by_day_3_4", lineSegments[lineSegmentCounter + 76].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 76].ToString()),
                                                new SqlParameter("@discount_by_day_3_5", lineSegments[lineSegmentCounter + 77].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 77].ToString()),
                                                new SqlParameter("@discount_by_day_3_6", lineSegments[lineSegmentCounter + 78].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 78].ToString()),
                                                new SqlParameter("@discount_by_day_3_7", lineSegments[lineSegmentCounter + 79].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 79].ToString()),
                                                new SqlParameter("@subscription_rate_description_4", lineSegments[lineSegmentCounter + 80].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 80].ToString()),
                                                new SqlParameter("@subscription_rate_before_tax_4", lineSegments[lineSegmentCounter + 81].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 81].ToString()),
                                                new SqlParameter("@subscription_rate_4", lineSegments[lineSegmentCounter + 82].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 82].ToString()),
                                                new SqlParameter("@new_expire_date_4", lineSegments[lineSegmentCounter + 83].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 83].ToString()),
                                                new SqlParameter("@discount_amount_4", lineSegments[lineSegmentCounter + 84].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 84].ToString()),
                                                new SqlParameter("@rate_by_day_4_1", lineSegments[lineSegmentCounter + 85].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 85].ToString()),
                                                new SqlParameter("@rate_by_day_4_2", lineSegments[lineSegmentCounter + 86].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 86].ToString()),
                                                new SqlParameter("@rate_by_day_4_3", lineSegments[lineSegmentCounter + 87].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 87].ToString()),
                                                new SqlParameter("@rate_by_day_4_4", lineSegments[lineSegmentCounter + 88].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 88].ToString()),
                                                new SqlParameter("@rate_by_day_4_5", lineSegments[lineSegmentCounter + 89].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 89].ToString()),
                                                new SqlParameter("@rate_by_day_4_6", lineSegments[lineSegmentCounter + 90].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 90].ToString()),
                                                new SqlParameter("@rate_by_day_4_7", lineSegments[lineSegmentCounter + 91].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 91].ToString()),
                                                new SqlParameter("@cost_by_day_4_1", lineSegments[lineSegmentCounter + 92].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 92].ToString()),
                                                new SqlParameter("@cost_by_day_4_2", lineSegments[lineSegmentCounter + 93].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 93].ToString()),
                                                new SqlParameter("@cost_by_day_4_3", lineSegments[lineSegmentCounter + 94].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 94].ToString()),
                                                new SqlParameter("@cost_by_day_4_4", lineSegments[lineSegmentCounter + 95].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 95].ToString()),
                                                new SqlParameter("@cost_by_day_4_5", lineSegments[lineSegmentCounter + 96].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 96].ToString()),
                                                new SqlParameter("@cost_by_day_4_6", lineSegments[lineSegmentCounter + 97].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 97].ToString()),
                                                new SqlParameter("@cost_by_day_4_7", lineSegments[lineSegmentCounter + 98].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 98].ToString()),
                                                new SqlParameter("@discount_by_day_4_1", lineSegments[lineSegmentCounter + 99].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 99].ToString()),
                                                new SqlParameter("@discount_by_day_4_2", lineSegments[lineSegmentCounter + 100].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 100].ToString()),
                                                new SqlParameter("@discount_by_day_4_3", lineSegments[lineSegmentCounter + 101].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 101].ToString()),
                                                new SqlParameter("@discount_by_day_4_4", lineSegments[lineSegmentCounter + 102].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 102].ToString()),
                                                new SqlParameter("@discount_by_day_4_5", lineSegments[lineSegmentCounter + 103].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 103].ToString()),
                                                new SqlParameter("@discount_by_day_4_6", lineSegments[lineSegmentCounter + 104].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 104].ToString()),
                                                new SqlParameter("@discount_by_day_4_7", lineSegments[lineSegmentCounter + 105].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 105].ToString()),
                                                new SqlParameter("@grace_owed_amount_not_including_taxes", lineSegments[lineSegmentCounter + 106].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 106].ToString()),
                                                new SqlParameter("@grace_owed_amount_including_taxes", lineSegments[lineSegmentCounter + 107].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 107].ToString()),
                                                new SqlParameter("@grace_owed_city_tax", lineSegments[lineSegmentCounter + 108].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 108].ToString()),
                                                new SqlParameter("@grace_owed_county_tax", lineSegments[lineSegmentCounter + 109].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 109].ToString()),
                                                new SqlParameter("@grace_owed_state_tax", lineSegments[lineSegmentCounter + 110].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 110].ToString()),
                                                new SqlParameter("@grace_owed_country_tax", lineSegments[lineSegmentCounter + 111].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 111].ToString()),
                                                new SqlParameter("@subscribers_unallocated_balance", lineSegments[lineSegmentCounter +112].ToString() == "" ? (object)DBNull.Value : lineSegments[lineSegmentCounter + 112].ToString()));
                                break;
                        }

                      

                    }



                }

            }

            //WriteToJobLog(JobLogMessageType.INFO, $"{routeDetailCounter + advanceDetailCounter + advanceTotalCounter + TMDetailCounter + TMTotalCounter + truckTotalCounter} total records read for publishing date {runDate.Value.ToShortDateString() ?? ""} type {runType}");
            //WriteToJobLog(JobLogMessageType.INFO, $"{routeDetailCounter} route detail read.");
            //WriteToJobLog(JobLogMessageType.INFO, $"{advanceDetailCounter} advance detail read.");
            //WriteToJobLog(JobLogMessageType.INFO, $"{advanceTotalCounter} advance total read.");
            //WriteToJobLog(JobLogMessageType.INFO, $"{TMDetailCounter} TM product detail read.");
            //WriteToJobLog(JobLogMessageType.INFO, $"{TMTotalCounter} TM product totals read.");
            //WriteToJobLog(JobLogMessageType.INFO, $"{truckTotalCounter} truck totals read.");

            //LoadRelatedTables(loadsId, runDate, runType, routeDetailCounter + advanceDetailCounter + advanceTotalCounter + TMDetailCounter + TMTotalCounter + truckTotalCounter);

        }

        public override void SetupJob()
        {
            JobName = "Office Pay";
            JobDescription = @"";
            AppConfigSectionName = "OfficePay";
        }
    }
}
