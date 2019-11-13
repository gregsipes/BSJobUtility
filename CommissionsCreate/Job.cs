using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommissionsCreate
{
    public class Job : JobBase
    {
        private enum CommissionCreateTypes
        {
            Create,
            RecreateForStructure,
            RecreateForSalesperson
        }

        public override void ExecuteJob()
        {
            Int64 commissionsId = -1;
            CommissionRecord commissionRecord = null;
            CommissionCreateTypes createType = CommissionCreateTypes.RecreateForSalesperson;


            using (var comm = new SqlCommand())
            {
                try
                {
                    // Checks the CommissionsCreate_Requested table for a record with a null session_uid and add set the session_uid if on is found.
                    //This record gets created by the Commissions interface
                    using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Update_CommissionsCreate_Requested", null))
                    {
                        if (!reader.HasRows)
                        {
                            WriteToJobLog(JobLogMessageType.INFO, "No commissions create requests exist");
                            return;
                        }

                        //set the commissions id
                        commissionsId = reader.GetInt64(reader.GetOrdinal("commissionscreate_requested_id"));

                        //build log mesage
                        string message = "Processing commissions create request by " + reader.GetString(reader.GetOrdinal("requested_user_name")) + " on " +
                                     String.Format("{0:MM/dd/yyyy hh:mm tt}", reader.GetDateTime(reader.GetOrdinal("requested_date_time")));

                        //todo: do we need the emailsubset process?

                        WriteToJobLog(JobLogMessageType.INFO, message);

                        int month = reader.GetInt32(reader.GetOrdinal("commissions_month"));
                        int year = reader.GetInt32(reader.GetOrdinal("commission_year"));
                        Int64 salespersonGroupId = -1;

                        if (reader.GetBoolean(reader.GetOrdinal("new_commissions_flag")))
                        {
                            //this is a new commissions run
                            createType = CommissionCreateTypes.Create;
                            commissionsId = -1;
                        }
                        else if (String.IsNullOrEmpty(reader.GetString(reader.GetOrdinal("salespersons_groups_id"))))
                        {
                            //this is a recreate for structure request
                            createType = CommissionCreateTypes.RecreateForStructure;
                        }
                        else
                        {
                            //this is a recreate for salesperson request
                            createType = CommissionCreateTypes.RecreateForSalesperson;
                            salespersonGroupId = reader.GetInt64(reader.GetOrdinal("salespersons_groups_id"));
                        }

                        //create commissions object
                        commissionRecord = new CommissionRecord() { Month = month, Year = year, CommissionsId = commissionsId };
                        commissionRecord.EndDate = reader.GetDateTime(reader.GetOrdinal("commissions_end_date"));
                        commissionRecord.MonthStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_month_start_date"));
                        commissionRecord.PriorEndDate = reader.GetDateTime(reader.GetOrdinal("commissions_prior_end_date"));
                        commissionRecord.PriorMonthStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_prior_month_start_date"));
                        commissionRecord.PriorYearStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_prior_ytd_start_date"));
                        commissionRecord.YearStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_ytd_start_date"));
                        commissionRecord.GainsLossesTopCount = reader.GetString(reader.GetOrdinal("gains_losses_top_count"));
                        commissionRecord.StructuresId = reader.GetInt64(reader.GetOrdinal("structures_id"));
                        commissionRecord.RequestedUserName = reader.GetString(reader.GetOrdinal("requested_user_name"));
                        commissionRecord.SalespersonName = reader.GetString(reader.GetOrdinal("salesperson_name"));
                        commissionRecord.SalespersonGroupId = reader.GetInt32(reader.GetOrdinal("salesperson_groups_id"));
                    }

                    //process commission request
                    ProcessCommissions(createType, commissionRecord);

                    //todo: build and send email



                    //delete request
                    ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Delete_CommissionsCreate_Requested",
                                            new Dictionary<string, object>()
                                            {
                                                { "@pintCommissionsCreateRequestedID", commissionsId }
                                            });
                }
                catch (Exception)
                {
                    throw;
                }
                finally
                {
                    if (comm != null && comm.Connection != null)
                        comm.Connection.Close();
                }
            }
        }

        public override void SetupJob()
        {
            JobName = "Commissions Create";
            JobDescription = "Creates monthly employee commission statements";
            AppConfigSectionName = "ParkingPayroll";
        }

        private void ProcessCommissions(CommissionCreateTypes createType, CommissionRecord commissionsRecord)
        {
            if (createType == CommissionCreateTypes.Create)
                CreateNewCommission(commissionsRecord); //new commissions create request
            else
                RecreateCommission(createType, commissionsRecord);   //recreate a commissions request



            DeleteAutoAttachments();
        }

        private void CreateNewCommission(CommissionRecord commissionsRecord)
        {
            WriteToJobLog(JobLogMessageType.INFO, "New commissions for " + commissionsRecord.StructuresId.ToString() + " " + commissionsRecord.Month.ToString() + "/" + commissionsRecord.Year);

            //Inserts a new record in the Commissions table and returns a new commissionId (unique value for this run)
            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions",
                                                new Dictionary<string, object>()
                                                {
                                                    { "@pintStructuresID", commissionsRecord.StructuresId },
                                                    { "@pintCommissionsYear", commissionsRecord.Year },
                                                    { "@pintCommissionsMonth", commissionsRecord.Month },
                                                    { "@psdatCommissionsYTDStartDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.YearStartDate) },
                                                    { "@psdatCommissionsMonthStartDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.MonthStartDate) },
                                                    { "@psdatCommissionsEndDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.EndDate) }
                                                }))
            {
                commissionsRecord.SpreadsheetStyle = reader.GetInt32(reader.GetOrdinal("spreadsheet_style"));
                //commissionsRecord.CommissionsId = reader.GetInt64(reader.GetOrdinal("commissions_id"));
                commissionsRecord.SnapshotId = reader.GetInt64(reader.GetOrdinal("snapshots_id"));
                commissionsRecord.PerformanceForBARCInsertStoredProcedure = reader.GetString(reader.GetOrdinal("performance_for_barc_insert_stored_procedure"));
                commissionsRecord.PlaybookForBARCInsertStoredProcedure = reader.GetString(reader.GetOrdinal("playbook_for_barc_insert_stored_procedure"));
                commissionsRecord.PlaybookForBARCUpdateStoredProcedure = reader.GetString(reader.GetOrdinal("playbook_for_barc_update_stored_procedure"));

            }

            if (GenerateCommissions(CommissionCreateTypes.Create,commissionsRecord))
                WriteToJobLog(JobLogMessageType.INFO, "Commissions created successfully");
            else
                WriteToJobLog(JobLogMessageType.WARNING, "Commissions could not be created");


        }

        private void RecreateCommission(CommissionCreateTypes createType, CommissionRecord commissionsRecord)
        {

            String message = "";
            if (createType == CommissionCreateTypes.RecreateForSalesperson)
                message = "Recreate commissions for " + commissionsRecord.SalespersonName + " in " + commissionsRecord.StructuresId.ToString() + " " +
                        commissionsRecord.Month.ToString() + "/" + commissionsRecord.Year.ToString();
            else
                message = "Recreate commissions for " + commissionsRecord.StructuresId.ToString() + " " +
                          commissionsRecord.Month.ToString() + "/" + commissionsRecord.Year.ToString();

            //todo: insert email subset

            WriteToJobLog(JobLogMessageType.INFO, message);

            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Commissions_Recreate",
                                                        new Dictionary<string, object>()
                                                        {
                                                            { "@pintStructuresID", commissionsRecord.StructuresId },
                                                            { "@pintCommissionsYear", commissionsRecord.Year },
                                                            { "@pintCommissionsMonth", commissionsRecord.Month }
                                                        }))
            {
                if (ValidateProcedure(reader, "Commissions cannot be recreated because other commissions are currently being recreated for this structure"))
                    return;
            }

            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Commissions_Paid_Processing",
                                            new Dictionary<string, object>()
                                            {
                                                            { "@pintStructuresID", commissionsRecord.StructuresId },
                                                            { "@pintCommissionsYear", commissionsRecord.Year },
                                                            { "@pintCommissionsMonth", commissionsRecord.Month }
                                            }))
            {
                if (ValidateProcedure(reader, "Commissions cannot be recreated because they are in the process of being paid by Payroll"))
                    return;
            }

            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Structures",
                                            new Dictionary<string, object>()
                                            {
                                                            { "@pintStructuresID", commissionsRecord.StructuresId }
                                            }))
            {
                if (!reader.GetBoolean(reader.GetOrdinal("verified_flag")))
                {
                    WriteToJobLog(JobLogMessageType.WARNING, "Structure (" + commissionsRecord.StructuresId + ") must be verified before salesperson's commissions can be recreated");
                    return;
                }
            }

            if (createType == CommissionCreateTypes.RecreateForSalesperson)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Statuses_Creating",
                                            new Dictionary<string, object>()
                                            {
                                                { "@pintStructuresID", commissionsRecord.StructuresId },
                                                { "@pintSalespersonsGroupsID", commissionsRecord.SalespersonGroupId },
                                                { "@pvchrSalespersonName", commissionsRecord.SalespersonName },
                                                { "@pvchrStatusBy", commissionsRecord.RequestedUserName }
                                            });
            }
            else
            {
                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_For_Commissions_Recreate",
                                                              new Dictionary<string, object>()
                                                              {
                                                                                { "@pintCommissionsID", commissionsRecord.CommissionsId },
                                                                                { "@pvchrUserName", commissionsRecord.RequestedUserName }
                                                              }))
                {
                    if (!reader.GetBoolean(reader.GetOrdinal("creating_flag")))
                    {
                        WriteToJobLog(JobLogMessageType.WARNING, "Recreate not creating");
                        return;
                    }

                    commissionsRecord.SnapshotId = reader.GetInt64(reader.GetOrdinal("snapshots_id"));
                }
            }


            if (GenerateCommissions(createType, commissionsRecord))
                WriteToJobLog(JobLogMessageType.INFO, "Commissions created successfully");
            else
                WriteToJobLog(JobLogMessageType.WARNING, "Commissions could not be created");

        }

        private bool GenerateCommissions(CommissionCreateTypes createType, CommissionRecord commissionRecord)
        {
            string commissionsFolder = GetConfigurationKeyValue("AttachmentDirectory");
            DateTime BARCDatetime;
            Int64 commissionsRecreateId = 0;

            WriteToJobLog(JobLogMessageType.INFO, "Initializing commissions");

            //ResetForExcel - is this needed?

            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.BuffNewsForBW, CommandType.StoredProcedure, "dbo.Proc_Select_Commissions_BuffNews_BARC_Populated", null))
            {
                if (!reader.HasRows)
                {
                    WriteToJobLog(JobLogMessageType.WARNING, "No BARC data is available for selection");
                    return false;
                }
                else
                    BARCDatetime = reader.GetDateTime(reader.GetOrdinal("end_date_time"));
            }


            if (createType != CommissionCreateTypes.Create)
            {

                //build commission object
                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Commissions_Related",
                                                    new Dictionary<string, object>()
                                                    {
                                                        { "@pintCommissionsID", commissionRecord.CommissionsId }
                                                    }))
                {
                    commissionRecord.EndDate = reader.GetDateTime(reader.GetOrdinal("commissions_end_date"));
                    commissionRecord.MonthStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_month_start_date"));
                    commissionRecord.PriorEndDate = reader.GetDateTime(reader.GetOrdinal("commissions_prior_end_date"));
                    commissionRecord.PriorMonthStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_prior_month_start_date"));
                    commissionRecord.PriorYearStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_prior_ytd_start_date"));
                    commissionRecord.YearStartDate = reader.GetDateTime(reader.GetOrdinal("commissions_ytd_start_date"));
                    commissionRecord.Month = reader.GetInt32(reader.GetOrdinal("commissions_month"));
                    commissionRecord.Year = reader.GetInt32(reader.GetOrdinal("commissions_year"));

                    commissionRecord.GainsLossesTopCount = reader.GetString(reader.GetOrdinal("gains_losses_top_count"));
                    commissionRecord.SpreadsheetStyle = reader.GetInt32(reader.GetOrdinal("spreadsheet_style"));
                    commissionRecord.StructuresId = reader.GetInt64(reader.GetOrdinal("structures_id"));
                    commissionRecord.PerformanceForBARCInsertStoredProcedure = reader.GetString(reader.GetOrdinal("performance_for_barc_insert_stored_procedure"));
                    commissionRecord.PlaybookForBARCInsertStoredProcedure = reader.GetString(reader.GetOrdinal("playbook_for_barc_insert_stored_procedure"));
                    commissionRecord.PlaybookForBARCUpdateStoredProcedure = reader.GetString(reader.GetOrdinal("playbook_for_barc_update_stored_procedure"));
                }


            }

            if (createType == CommissionCreateTypes.RecreateForSalesperson)
            {
                //set snapshot id (unique id for the run)
                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots", null))
                {
                    commissionRecord.SnapshotId = reader.GetInt64(reader.GetOrdinal("snapshots_id"));
                }


                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Recreate",
                                                                    new Dictionary<string, object>()
                                                                    {
                                                                        { "@pintStructuresID", commissionRecord.StructuresId },
                                                                        { "@pintCommissionsYear", commissionRecord.Year },
                                                                        { "@pintCommissionsMonth", commissionRecord.Month },
                                                                        { "@psdatCommissionYTDStartDate", commissionRecord.YearStartDate },
                                                                        { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                                        { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                                                        { "@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId },
                                                                        { "@pintNewSnapshotsID", commissionRecord.SnapshotId },
                                                                        { "@pvchrRecreateBy", commissionRecord.RequestedUserName },
                                                                        { "@pvchrRecreateComputerName", "" }
                                                                    }))
                {
                    string message = reader.GetString(reader.GetOrdinal("message"));

                    if (!String.IsNullOrEmpty(message))
                    {
                        WriteToJobLog(JobLogMessageType.WARNING, message);
                        return false;
                    }

                    commissionsRecreateId = reader.GetInt64(reader.GetOrdinal("commissions_recreate_id"));
                }

                //take a snapshot of each table
                TakeSnapshot(commissionsRecreateId, "BARC");
                TakeSnapshot(commissionsRecreateId, "Data_Mining");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Accounts");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Chargebacks");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Draw_Per_Days");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Noncommissions");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Nonworking_Dates");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Playbook_Groups");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Playbook_Print_Division_Descriptions");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Product_Data_Mining_Descriptions");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Product_Groups");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Responsible_Salespersons");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Salespersons");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Salespersons_Groups");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Strategies");
                TakeSnapshot(commissionsRecreateId, "Snapshots_Territories");

            }


            //get salespersons
            Dictionary<string, string> salespersons = new Dictionary<string, string>();
            if (createType == CommissionCreateTypes.Create | createType == CommissionCreateTypes.RecreateForStructure)
            {
                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Salespersons",
                                                        new Dictionary<string, object>()
                                                        {
                                                            { "@pintSnapshotsID", commissionRecord.SnapshotId }
                                                        }))
                {
                    while (reader.Read())
                    {
                        salespersons.Add(reader.GetString(reader.GetOrdinal("salesperson")), reader.GetString(reader.GetOrdinal("salesperson_name")));
                    }
                }
            }
            else
            {
                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Salespersons_Recreate",
                                                            new Dictionary<string, object>()
                                                            {
                                                                { "@pintCommissionsRecreateID", commissionRecord.CommissionsId }
                                                             }))
                {
                    while (reader.Read())
                    {
                        salespersons.Add(reader.GetString(reader.GetOrdinal("salesperson")), reader.GetString(reader.GetOrdinal("salesperson_name")));
                    }
                }
            }

            //get commissions inquiry id
            Int64 commissionsInquiriesId = 0;
            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.CommissionsRelated, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Inquiries",
                                                                new Dictionary<string, object>()
                                                                {
                                                                    { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                                                    { "@pintCommissionsYear",commissionRecord.Year },
                                                                    { "@pintCommissionsMonth", commissionRecord.Month },
                                                                    { "@psdatCommissionsYTDStartDate", commissionRecord.YearStartDate },
                                                                    { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                                    { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                                                    { "@psdatCommissionsPriorYTDStartDate", commissionRecord.PriorYearStartDate },
                                                                    { "@psdatCommissionsPriorMonthStartDate", commissionRecord.PriorMonthStartDate },
                                                                    { "@psdatCommissionsPriorEndDate", commissionRecord.PriorEndDate },
                                                                    { "@pintGainsLossesTopCount", commissionRecord.GainsLossesTopCount }
                                                                }))
            {
                commissionsInquiriesId = reader.GetInt64(reader.GetOrdinal("commissions_inquiries_id"));
            }

            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Product_Data_Mining_Descriptions",
                                                    new Dictionary<string, object>()
                                                    {
                                                                    { "@pintCommissionsInquiriesID", commissionsInquiriesId }
                                                    }))
            {
                while (reader.Read())
                {
                    ExecuteNonQuery(DatabaseConnectionStringNames.CommissionsRelated, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Inquiries_Data_Mining",
                                                    new Dictionary<string, object>()
                                                    {
                                                         { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                         {  "@pvchrtblEditionsEdnNumber", reader.GetString(reader.GetOrdinal("tbleditions_ednnumber")) }
                                                    });
                }
            }


            ExecuteNonQuery(DatabaseConnectionStringNames.BuffNewsForBW, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Inquiries",
                                                    new Dictionary<string, object>()
                                                    {
                                                        { "@pvchrCommissionsRelatedServerInstance", GetConfigurationKeyValue("CommissionsRelatedServerName") },
                                                        { "@pvchrCommissionsRelatedDatabase", GetConfigurationKeyValue("CommissionsRelatedDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId }
                                                    });

            foreach (var salesperson in salespersons)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.CommissionsRelated, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Inquiries_Responsible_Salespersons",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrSalesperson", salesperson.Key }
                                        });
                ExecuteNonQuery(DatabaseConnectionStringNames.CommissionsRelated, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Inquiries_Performance_Salespersons",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrSalesperson", salesperson.Key }
                                        });
            }


            ExecuteNonQuery(DatabaseConnectionStringNames.CommissionsRelated, CommandType.StoredProcedure, "dbo.Proc_Insert_Commissions_Inquiries_Responsible_Salespersons",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pvchrCommissionsRelatedServerInstance", GetConfigurationKeyValue("CommissionsRelatedServerName") },
                                                        { "@pvchrCommissionsRelatedDatabase", GetConfigurationKeyValue("CommissionsRelatedDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId }
                                        });

            WriteToJobLog(JobLogMessageType.INFO, "Selecting menu mania data mining data from Brainworks");

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Data_Mining",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pvchrBrainworksServerInstance", GetConfigurationKeyValue("BrainworksServerName") },
                                                        { "@pvchrBrainworksRelatedDatabase", GetConfigurationKeyValue("BrainworksDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrStoredProcedure", "Proc_BuffNews_Select_Commissions_Data_Mining_Menu_Mania" }
                                        });

            WriteToJobLog(JobLogMessageType.INFO, "Selecting new business data mining data from Brainworks");

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Data_Mining",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pvchrBrainworksServerInstance", GetConfigurationKeyValue("BrainworksServerName") },
                                                        { "@pvchrBrainworksRelatedDatabase", GetConfigurationKeyValue("BrainworksDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrStoredProcedure", "Proc_BuffNews_Select_Commissions_Data_Mining_New_Business" }
                                        });

            WriteToJobLog(JobLogMessageType.INFO, "Selecting product data mining data from Brainworks");

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Data_Mining",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pvchrBrainworksServerInstance", GetConfigurationKeyValue("BrainworksServerName") },
                                                        { "@pvchrBrainworksRelatedDatabase", GetConfigurationKeyValue("BrainworksDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrStoredProcedure", "Proc_BuffNews_Select_Commissions_Data_Mining_Product" }
                                        });

            WriteToJobLog(JobLogMessageType.INFO, "Selecting playbook data from BARC");
            //this is pulling in a snapshot of the BuffNewsForBW.BuffNews_BARC_Brainworks table depending which sproc is passed in
            //Does not create any new records
            //'Proc_Insert_BARC “BWDB\BW,50884', 'BuffNewsForBW', 'CommissionsCreate', '<Cr#@t0rUs3r>', 2607, 'Proc_Select_Commissions_Outside_Auto_Playbook_Detail'
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_BARC",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName") },
                                                        { "@pvchrBuffNewsForBWRelatedDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrStoredProcedure", commissionRecord.PlaybookForBARCInsertStoredProcedure }
                                        });

            WriteToJobLog(JobLogMessageType.INFO, "Selecting performance data from BARC");
            //Does not create any new records
            //Proc_Insert_BARC “BWDB\BW,50884', 'BuffNewsForBW', 'CommissionsCreate', '<Cr#@t0rUs3r>', 2607, 'Proc_Select_Commissions_Outside_Auto_Performance_Detail'
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_BARC",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName") },
                                                        { "@pvchrBuffNewsForBWRelatedDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrStoredProcedure", commissionRecord.PerformanceForBARCInsertStoredProcedure }
                                        });

            WriteToJobLog(JobLogMessageType.INFO, "Selecting gains/losses data from BARC");
            //Creates 631 new records with new snapshots_id.  HOW DID THE SNAPSHOTS ID GET INTO HERE???????
            //Proc_Insert_BARC “BWDB\BW,50884', 'BuffNewsForBW', 'CommissionsCreate', '<Cr#@t0rUs3r>', 2607, 'Proc_Select_Commissions_Gains_Losses_Detail ‘”
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_BARC",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName") },
                                                        { "@pvchrBuffNewsForBWRelatedDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName") },
                                                        { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                                        { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword") },
                                                        { "@pintCommissionsInquiriesID", commissionsInquiriesId },
                                                        { "@pvchrStoredProcedure", "Proc_Select_Commissions_Gains_Losses_Detail" }
                                        });


            WriteToJobLog(JobLogMessageType.INFO, "Initializing snapshots");
            RunSnapshotSprocs(commissionRecord, createType, commissionsRecreateId, commissionRecord.SnapshotId, salespersons);




            return true;

        }

        private void RunSnapshotSprocs(CommissionRecord commissionRecord, CommissionCreateTypes createType, Int64 commissionsRecreateId, Int64 snapshotId, Dictionary<string, string> salespersons)
        {
            //only execute if we are recreating for a salesperson
            if (createType == CommissionCreateTypes.RecreateForSalesperson)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Territories",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pintStructuresID", commissionRecord.StructuresId }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Salespersons_Groups",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Salespersons",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId }
                                        });
            }

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Accounts",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pintCommissionsYear", commissionRecord.Year },
                                                         { "@pintCommissionsMonth", commissionRecord.Month }
                                        });

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Noncommissions",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pintCommissionsYear", commissionRecord.Year },
                                                         { "@pintCommissionsMonth", commissionRecord.Month }
                                        });

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Chargebacks",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pintCommissionsYear", commissionRecord.Year },
                                                         { "@pintCommissionsMonth", commissionRecord.Month }
                                        });

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Nonworking_Dates",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                         { "@psdatCommissionsEndDate", commissionRecord.EndDate }
                                        });

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Draw_Per_Days",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                         { "@psdatCommissionsEndDate", commissionRecord.EndDate }
                                        });

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Update_Snapshots_Salespersons_Performance_Goal_Percentage",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pintCommissionsYear", commissionRecord.Year },
                                                         { "@pintCommissionsMonth", commissionRecord.Month }
                                         });

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Strategies",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintCommissionsRecreateID", commissionsRecreateId },
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pintCommissionsYear", commissionRecord.Year },
                                                         { "@pintCommissionsMonth", commissionRecord.Month }
                                        });

            foreach (var salesperson in salespersons)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Playbook_Groups",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pvchrSalesperson", salesperson.Key },
                                                         { "@pintCommissionsYear", commissionRecord.Year },
                                                         { "@pintCommissionsMonth", commissionRecord.Month }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Playbook_Print_Division_Descriptions",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pvchrSalesperson", salesperson.Key }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, commissionRecord.PlaybookForBARCUpdateStoredProcedure,
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pintSnapshotsID", snapshotId },
                                                        { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                        { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                                        { "@pvchrSalesperson", salesperson.Key }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Product_Groups",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pvchrSalesperson", salesperson.Key },
                                                         { "@pintCommissionsYear", commissionRecord.Year },
                                                         { "@pintCommissionsMonth", commissionRecord.Month }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Snapshots_Product_Data_Mining_Descriptions",
                                        new Dictionary<string, object>()
                                        {
                                                         { "@pintSnapshotsID", snapshotId },
                                                         { "@pvchrSalesperson", salesperson.Key }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Update_Snapshots_Product_Groups_Product",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pintSnapshotsID", snapshotId },
                                                        { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                        { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                                        { "@pvchrSalesperson", salesperson.Key }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Update_Snapshots_Product_Groups_Menu_Mania",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pintSnapshotsID", snapshotId },
                                                        { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                        { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                                        { "@pvchrSalesperson", salesperson.Key }
                                        });

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Update_Snapshots_Product_Groups_New_Business",
                                        new Dictionary<string, object>()
                                        {
                                                        { "@pintSnapshotsID", snapshotId },
                                                        { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                                        { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                                        { "@pvchrSalesperson", salesperson.Key }
                                        });
            }

        }

        private void TakeSnapshot(Int64 commissisionRecreateId, string tableName)
        {
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Copy_Between_Snapshots",
                                        new Dictionary<string, object>()
                                        {
                                                            { "@pintCommissionsRecreateID", commissisionRecreateId },
                                                            { "@pvchrTableName", tableName }
                                        });
        }

        /// <summary>
        /// Validate the execute of a stored procedure that run during the recreate commmission process
        /// </summary>
        /// <param name="comm">Command to be executed</param>
        /// <param name="message">Log message prefix</param>
        /// <returns></returns>
        private bool ValidateProcedure(SqlDataReader reader, string message)
        {
            if (reader.HasRows)
            {
                WriteToJobLog(JobLogMessageType.WARNING, message + " by " + reader.GetString(reader.GetOrdinal("processing_by")) + " at " +
                                    String.Format("{0:MM/dd/yyyy hh:mm tt}", reader.GetDateTime(reader.GetOrdinal("processing_date_time"))));
                return false;
            }

            return true;
        }

    }
}

