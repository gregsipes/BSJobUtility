using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ManifestLoadAdvance
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Manifest Load Advance";
            JobDescription = "Builds advance manifest labels";
            AppConfigSectionName = "ManifestLoadAdvance";
        }

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "advmanifest*").ToList();


                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Select_Loads_If_Processed",
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
                SendMail($"Error in Job: {JobName}", ex.ToString(), false);
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
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Loads",
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

            Int32 routeDetailCounter = 0;
            Int32 advanceDetailCounter = 0;
            Int32 advanceTotalCounter = 0;
            Int32 TMDetailCounter = 0;
            Int32 TMTotalCounter = 0;
            Int32 truckTotalCounter = 0;

            DateTime? runDate = null;
            String runType = "";

            foreach (string line in fileContents)
            {
                if (line != null && line.Trim().Length > 0)
                {
                    List<string> lineSegments = line.Split('|').ToList();

                    if (lineSegments[0] == "R1")
                    {
                        routeDetailCounter++;

                        if (lineSegments.Count() < 41)
                            WriteToJobLog(JobLogMessageType.ERROR, $"Error on record # {routeDetailCounter} . Field count: {lineSegments.Count()}  Line: {line}");
                        else
                        {

                            runDate = DateTime.Parse(lineSegments[4].ToString());
                            runType = lineSegments[3].ToString();

                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Route_Detail",
                                            new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@record_number", routeDetailCounter),
                                            new SqlParameter("@record_type", lineSegments[0].ToString()),
                                            new SqlParameter("@main_product_id", lineSegments[1].ToString()),
                                            new SqlParameter("@product_id", lineSegments[2].ToString()),
                                            new SqlParameter("@run_type", lineSegments[3].ToString()),
                                            new SqlParameter("@run_date", lineSegments[4].ToString()),
                                            new SqlParameter("@truck_id", lineSegments[5].ToString()),
                                            new SqlParameter("@truck_name", lineSegments[6].ToString()),
                                            new SqlParameter("@truck_drop_order", lineSegments[7].ToString()),
                                            new SqlParameter("@route_id_or_relay_truck_id", lineSegments[8].ToString()),
                                            new SqlParameter("@route_type_or_single_copy_type", lineSegments[9].ToString()),
                                            new SqlParameter("@route_type_indicator", lineSegments[10].ToString()),
                                            new SqlParameter("@depot_id", lineSegments[11].ToString()),
                                            new SqlParameter("@depot_drop_order", lineSegments[12].ToString()),
                                            new SqlParameter("@edition_or_paper_section", lineSegments[13].ToString()),
                                            new SqlParameter("@draw_total", lineSegments[14].ToString()),
                                            new SqlParameter("@number_of_standard_bundles", lineSegments[15].ToString()),
                                            new SqlParameter("@number_of_key_bundles", lineSegments[16].ToString()),
                                            new SqlParameter("@key_bundle_size", lineSegments[17].ToString()),
                                            new SqlParameter("@carrier_name", lineSegments[18].ToString()),
                                            new SqlParameter("@carrier_phone_number", lineSegments[19].ToString()),
                                            new SqlParameter("@insert_mix_combination", lineSegments[20].ToString()),
                                            new SqlParameter("@drop_location", lineSegments[21].ToString()),
                                            new SqlParameter("@drop_instructions", lineSegments[22].ToString()),
                                            new SqlParameter("@ad_zone", lineSegments[23].ToString()),
                                            new SqlParameter("@preprint_demographic", lineSegments[24].ToString()),
                                            new SqlParameter("@insert_exception_indicator", lineSegments[25].ToString()),
                                            new SqlParameter("@bulk_indicator", lineSegments[26].ToString()),
                                            new SqlParameter("@hand_tie_indicator", lineSegments[27].ToString()),
                                            new SqlParameter("@minimum_bundle_size", lineSegments[28].ToString()),
                                            new SqlParameter("@maximum_bundle_size", lineSegments[29].ToString()),
                                            new SqlParameter("@standard_bundle_size", lineSegments[30].ToString()),
                                            new SqlParameter("@route_name_or_single_copy_location", lineSegments[31].ToString()),
                                            new SqlParameter("@map_reference", lineSegments[32].ToString()),
                                            new SqlParameter("@map_number", lineSegments[33].ToString()),
                                            new SqlParameter("@multipack_id", lineSegments[34].ToString()),
                                            new SqlParameter("@product_route_combination_weight", lineSegments[35].ToString()),
                                            new SqlParameter("@total_drop_weight", lineSegments[36].ToString()),
                                            new SqlParameter("@standard_bundle_weight", lineSegments[37].ToString()),
                                            new SqlParameter("@carrier_id", lineSegments[38].ToString()),
                                            new SqlParameter("@chute_number", lineSegments[39].ToString()),
                                            new SqlParameter("@departure_order", lineSegments[40].ToString()));
                        }
                    }
                    else if (lineSegments[0] == "R2")
                    {
                        advanceDetailCounter++;

                        if (lineSegments.Count() < 15)
                            WriteToJobLog(JobLogMessageType.ERROR, $"Error on record # {advanceDetailCounter} . Field count: {lineSegments.Count()}  Line: {line}");
                        else
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Advance_Detail",
                                        new SqlParameter("@loads_id", loadsId),
                                       new SqlParameter("@record_number", advanceDetailCounter),
                                       new SqlParameter("@record_type", lineSegments[0].ToString()),
                                       new SqlParameter("@main_product_id", lineSegments[1].ToString()),
                                       new SqlParameter("@product_id", lineSegments[2].ToString()),
                                       new SqlParameter("@run_type", lineSegments[3].ToString()),
                                       new SqlParameter("@run_date", lineSegments[4].ToString()),
                                       new SqlParameter("@truck_id", lineSegments[5].ToString()),
                                       new SqlParameter("@truck_name", lineSegments[6].ToString()),
                                       new SqlParameter("@truck_drop_order", lineSegments[7].ToString()),
                                       new SqlParameter("@route_id_or_relay_truck_id", lineSegments[8].ToString()),
                                       new SqlParameter("@paper_section", lineSegments[9].ToString()),
                                       new SqlParameter("@updraw", lineSegments[10].ToString()),
                                       new SqlParameter("@insert_mix_combination", lineSegments[11].ToString()),
                                       new SqlParameter("@carrier_id", lineSegments[12].ToString()),
                                       new SqlParameter("@chute_number", lineSegments[13].ToString()),
                                       new SqlParameter("@departure_order", lineSegments[14].ToString()));
                        }
                    }
                    else if (lineSegments[0] == "R3")
                    {
                        TMDetailCounter++;

                        if (lineSegments.Count() < 21)
                            WriteToJobLog(JobLogMessageType.ERROR, $"Error on record # {TMDetailCounter} . Field count: {lineSegments.Count()}  Line: {line}");
                        else
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_TM_Detail",
                                       new SqlParameter("@loads_id", loadsId),
                                       new SqlParameter("@record_number", TMDetailCounter),
                                       new SqlParameter("@record_type", TMDetailCounter),
                                       new SqlParameter("@main_product_id", lineSegments[0].ToString()),
                                       new SqlParameter("@product_id", lineSegments[1].ToString()),
                                       new SqlParameter("@run_type", lineSegments[2].ToString()),
                                       new SqlParameter("@run_date", lineSegments[3].ToString()),
                                       new SqlParameter("@truck_id", lineSegments[4].ToString()),
                                       new SqlParameter("@truck_name", lineSegments[5].ToString()),
                                       new SqlParameter("@truck_drop_order", lineSegments[6].ToString()),
                                       new SqlParameter("@route_id_or_relay_truck_id", lineSegments[7].ToString()),
                                       new SqlParameter("@tm_product_id", lineSegments[8].ToString()),
                                       new SqlParameter("@tm_draw_total", lineSegments[9].ToString()),
                                       new SqlParameter("@number_of_standard_bundles", lineSegments[10].ToString()),
                                       new SqlParameter("@number_of_key_bundles", lineSegments[11].ToString()),
                                       new SqlParameter("@totals_key_draw", lineSegments[12].ToString()),
                                       new SqlParameter("@minimum_bundle_size", lineSegments[13].ToString()),
                                       new SqlParameter("@maximum_bundle_size", lineSegments[14].ToString()),
                                       new SqlParameter("@standard_bundle_size", lineSegments[15].ToString()),
                                       new SqlParameter("@weight", lineSegments[16].ToString()),
                                       new SqlParameter("@standard_bundle_weight", lineSegments[17].ToString()),
                                       new SqlParameter("@carrier_id", lineSegments[18].ToString()),
                                       new SqlParameter("@chute_number", lineSegments[19].ToString()),
                                       new SqlParameter("@departure_order", lineSegments[20].ToString()));
                        }
                    }
                    else if (lineSegments[0] == "T1")
                    {
                        truckTotalCounter++;

                        if (lineSegments.Count() < 26)
                            WriteToJobLog(JobLogMessageType.ERROR, $"Error on record # {truckTotalCounter} . Field count: {lineSegments.Count()}  Line: {line}");
                        else
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Truck_Totals",
                                       new SqlParameter("@loads_id", loadsId),
                                       new SqlParameter("@record_number", truckTotalCounter),
                                       new SqlParameter("@record_type", lineSegments[0].ToString()),
                                       new SqlParameter("@main_product_id", lineSegments[1].ToString()),
                                       new SqlParameter("@product_id", lineSegments[2].ToString()),
                                       new SqlParameter("@run_type", lineSegments[3].ToString()),
                                       new SqlParameter("@run_date", lineSegments[4].ToString()),
                                       new SqlParameter("@truck_id", lineSegments[5].ToString()),
                                       new SqlParameter("@truck_name", lineSegments[6].ToString()),
                                       new SqlParameter("@insert_mix_combination", lineSegments[7].ToString()),
                                       new SqlParameter("@number_of_bundles", lineSegments[8].ToString()),
                                       new SqlParameter("@total_draw", lineSegments[9].ToString()),
                                       new SqlParameter("@key_draw", lineSegments[10].ToString()),
                                       new SqlParameter("@number_of_standard_bundles", lineSegments[11].ToString()),
                                       new SqlParameter("@number_of_key_bundles", lineSegments[12].ToString()),
                                       new SqlParameter("@bulk_draw_total", lineSegments[13].ToString()),
                                       new SqlParameter("@bulk_key_draw_total", lineSegments[14].ToString()),
                                       new SqlParameter("@number_of_bulk_standard_bundle_tops", lineSegments[15].ToString()),
                                       new SqlParameter("@number_of_bulk_key_bundle_tops", lineSegments[16].ToString()),
                                       new SqlParameter("@number_of_hand_ties", lineSegments[17].ToString()),
                                       new SqlParameter("@number_of_throwoffs", lineSegments[18].ToString()),
                                       new SqlParameter("@rounded_draw", lineSegments[19].ToString()),
                                       new SqlParameter("@total_weight", lineSegments[20].ToString()),
                                       new SqlParameter("@chute_number", lineSegments[21].ToString()),
                                       new SqlParameter("@drivers_name", lineSegments[22].ToString()),
                                       new SqlParameter("@departure_order", lineSegments[23].ToString()),
                                       new SqlParameter("@extra_1", lineSegments[24].ToString()),
                                       new SqlParameter("@extra_2", lineSegments[25].ToString()));
                        }
                    }
                    else if (lineSegments[0] == "T2")
                    {
                        advanceTotalCounter++;

                        if (lineSegments.Count() < 14)
                            WriteToJobLog(JobLogMessageType.ERROR, $"Error on record # {advanceDetailCounter} . Field count: {lineSegments.Count()}  Line: {line}");
                        else
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Advance_Totals",
                                       new SqlParameter("@loads_id", loadsId),
                                       new SqlParameter("@record_number", advanceTotalCounter),
                                       new SqlParameter("@record_type", lineSegments[0].ToString()),
                                       new SqlParameter("@main_product_id", lineSegments[1].ToString()),
                                       new SqlParameter("@product_id", lineSegments[2].ToString()),
                                       new SqlParameter("@run_type", lineSegments[3].ToString()),
                                       new SqlParameter("@run_date", lineSegments[4].ToString()),
                                       new SqlParameter("@truck_id", lineSegments[5].ToString()),
                                       new SqlParameter("@truck_name", lineSegments[6].ToString()),
                                       new SqlParameter("@paper_section", lineSegments[7].ToString()),
                                       new SqlParameter("@updraw", lineSegments[8].ToString()),
                                       new SqlParameter("@chute_number", lineSegments[9].ToString()),
                                       new SqlParameter("@drivers_name", lineSegments[10].ToString()),
                                       new SqlParameter("@departure_order", lineSegments[11].ToString()),
                                       new SqlParameter("@main_product_description", lineSegments[12].ToString()),
                                       new SqlParameter("@product_description", lineSegments[13].ToString()));
                        }
                    }
                    else if (lineSegments[0] == "T3")
                    {
                        TMTotalCounter++;

                        if (lineSegments.Count() < 21)
                            WriteToJobLog(JobLogMessageType.ERROR, $"Error on record # {TMTotalCounter} . Field count: {lineSegments.Count()}  Line: {line}");
                        else
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_TM_Totals",
                                            new SqlParameter("@loads_id", loadsId),
                                           new SqlParameter("@record_number", TMTotalCounter),
                                           new SqlParameter("@record_type", lineSegments[0].ToString()),
                                           new SqlParameter("@main_product_id", lineSegments[1].ToString()),
                                           new SqlParameter("@product_id", lineSegments[2].ToString()),
                                           new SqlParameter("@run_type", lineSegments[3].ToString()),
                                           new SqlParameter("@run_date", lineSegments[4].ToString()),
                                           new SqlParameter("@truck_id", lineSegments[5].ToString()),
                                           new SqlParameter("@truck_name", lineSegments[6].ToString()),
                                           new SqlParameter("@tm_product_id", lineSegments[7].ToString()),
                                           new SqlParameter("@total_drops", lineSegments[8].ToString()),
                                           new SqlParameter("@total_bundles", lineSegments[9].ToString()),
                                           new SqlParameter("@tm_draw_total", lineSegments[10].ToString()),
                                           new SqlParameter("@total_standard_bundles", lineSegments[11].ToString()),
                                           new SqlParameter("@total_key_bundles", lineSegments[12].ToString()),
                                           new SqlParameter("@tm_total_bulk_draw", lineSegments[13].ToString()),
                                           new SqlParameter("@total_bulk_standard_bundles", lineSegments[14].ToString()),
                                           new SqlParameter("@total_bulk_key_bundles", lineSegments[15].ToString()),
                                           new SqlParameter("@bulk_key_size", lineSegments[16].ToString()),
                                           new SqlParameter("@total_weight", lineSegments[17].ToString()),
                                           new SqlParameter("@chute_number", lineSegments[18].ToString()),
                                           new SqlParameter("@drivers_name", lineSegments[19].ToString()),
                                           new SqlParameter("@departure_order", lineSegments[20].ToString()));
                        }
                    }
                    //else
                    //{
                    //    WriteToJobLog(JobLogMessageType.ERROR, $"Error on line: {line}");
                    //    throw new Exception("File incorrectly formatted, exiting process");
                    //}
                }
            }

            WriteToJobLog(JobLogMessageType.INFO, $"{routeDetailCounter + advanceDetailCounter + advanceTotalCounter + TMDetailCounter + TMTotalCounter + truckTotalCounter} total records read for publishing date {runDate.Value.ToShortDateString() ?? ""} type {runType}");
            WriteToJobLog(JobLogMessageType.INFO, $"{routeDetailCounter} route detail read.");
            WriteToJobLog(JobLogMessageType.INFO, $"{advanceDetailCounter} advance detail read.");
            WriteToJobLog(JobLogMessageType.INFO, $"{advanceTotalCounter} advance total read.");
            WriteToJobLog(JobLogMessageType.INFO, $"{TMDetailCounter} TM product detail read.");
            WriteToJobLog(JobLogMessageType.INFO, $"{TMTotalCounter} TM product totals read.");
            WriteToJobLog(JobLogMessageType.INFO, $"{truckTotalCounter} truck totals read.");

            LoadRelatedTables(loadsId, runDate, runType, routeDetailCounter + advanceDetailCounter + advanceTotalCounter + TMDetailCounter + TMTotalCounter + truckTotalCounter);

        }

        private void LoadRelatedTables(Int32 loadsId, DateTime? runDate, String runType, Int32 totalRecordsProcessed)
        {
            WriteToJobLog(JobLogMessageType.INFO, $"Begin processing for load-related reference tables for loads_id {loadsId}.");

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Mother_Trucks_Sequence", new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, "Records inserted into Mother_Trucks_Sequence table.");

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Relay_Trucks_Sequence", new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, "Records inserted into Relay_Trucks_Sequence & Relay_Trucks_Sequence tables.");

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Mother_Trucks_Sequence_For_Loading_Dock", new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, "Records inserted into Editions_Sections_Sequence table.");

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Editions_Sections_No_AMPM_Sequence", new SqlParameter("@pintLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, "Records inserted into Editions_Sections_No_AMPM_Sequence table.");

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Insert_Loads_Latest",
                                                new SqlParameter("@pintLoadsID", loadsId),
                                                new SqlParameter("@psdatRunDate", runDate.Value.ToShortDateString() ?? null),
                                                new SqlParameter("@pvchrRunType", runType),
                                                new SqlParameter("@pintPageOrRecordCount", totalRecordsProcessed),
                                                new SqlParameter("@pflgSuccessful", 1));


              CheckInsertMixCombinationPrefix(runDate);
        }

        private void CheckInsertMixCombinationPrefix(DateTime? runDate)
        {
            WriteToJobLog(JobLogMessageType.INFO, "Checking for insert/mix combination used by more that one advance manifest.");

            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Select_Duplicate_Insert_Mix_Combination_Prefixes").ToList();

            if (results == null || results.Count() == 0)
                WriteToJobLog(JobLogMessageType.INFO, "No insert/mix combination prefix used by more than one advance manifest.");
            else
            {
                StringBuilder stringBuilder = new StringBuilder();

                stringBuilder.Append("Prefixes Used:");

                foreach (Dictionary<string, object> result in results)
                {
                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("\t\t Section: " + result["edition_or_paper_section"].ToString());
                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("\t\t\t Run Date: " + runDate.Value.ToShortDateString() ?? "");
                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("\t\t\t File: " + result["original_dir"].ToString() + result["original_file"].ToString());
                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("\t\t\t Last Modified: " + DateTime.Parse(result["original_file_last_modified"].ToString()).ToLongDateString());
                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("\t\t\t Loaded: " + DateTime.Parse(result["load_date"].ToString()).ToLongDateString());
                    stringBuilder.AppendLine();
                    stringBuilder.AppendLine("\t\t\t Loads ID: " + result["loads_id"].ToString());
                }

                SendMail("ManifestsLoad: Insert/Mix Combination Prefixes Used By More That One Advance Manifest", stringBuilder.ToString(), false);

            }



        }
    }

}
