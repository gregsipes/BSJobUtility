using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using System.Globalization;
using BSJobBase.Classes;

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

        private enum AutoAttachmentTypes
        {
            Playbook = 0,
            Products = 1,
            NewBusiness = 2,
            MenuMania = 3
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

            if (GenerateCommissions(CommissionCreateTypes.Create, commissionsRecord))
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

            bool isSuccessful = CreateCommissionsSpreadsheets(createType, commissionRecord);




            return true;

        }

        private bool CreateCommissionsSpeadsheets(CommissionCreateTypes createTypes, CommissionRecord commissionRecord)
        {
            //insert session
            Int64 sessionId = 0;
            using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Insert_Sessions",
                                                    new Dictionary<string, object>()
                                                    {
                                                          { "@pvchrUserName", commissionRecord.RequestedUserName }
                                                    }))
            {
                sessionId = reader.GetInt64(reader.GetOrdinal("sessions_id"));
            }


            //build salesperson groups
            List<SalespersonGroup> salespersonGroups = new List<SalespersonGroup>();

            if (createTypes == CommissionCreateTypes.Create)
            {
                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Snapshots_Salespersons_Groups",
                                        new Dictionary<string, object>()
                                        {
                                               { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                               { "@plngTerritoriesID", -1 }
                                        }))
                {
                    salespersonGroups = BuildSalespersonGroup(reader);
                }
            }
            else
            {
                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Snapshots_Salespersons_Groups_Recreate",
                        new Dictionary<string, object>()
                        {
                                               { "@pintCommissionsRecreateID", commissionRecord.CommissionsId }
                        }))
                {
                    salespersonGroups = BuildSalespersonGroup(reader);
                }
            }

            //iterate groups and create commissions statements for each
            foreach (SalespersonGroup salespersonGroup in salespersonGroups)
            {
                WriteToJobLog(JobLogMessageType.INFO, "Creating Commissions spreadsheet for " + salespersonGroup.SalespersonName);

                //Object excelApp = CreateObject("Excel.Application");
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Application.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Application.ActiveWorkbook;
                excel.Application.DisplayAlerts = false;

                Microsoft.Office.Interop.Excel.Worksheet activeWorksheet;
                //remove all worksheets except the first one
                while (workbook.Worksheets.Count > 1)
                {
                    activeWorksheet = workbook.Sheets[2];
                    activeWorksheet.Delete();
                }

                excel.Application.DisplayAlerts = true;

                using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Snapshots_Salespersons",
                                            new Dictionary<string, object>()
                                            {
                                               { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                               { "@pintSalespersonGroupsID", commissionRecord.SalespersonGroupId }
                                            }))
                {
                    bool isSummaryRecord = false;
                    while (true) //iterate salespersons 
                    {
                        string salesperson = "";
                        string salespersonGroup = "";
                        if (isSummaryRecord)
                            salesperson = "Summary For " + salespersonGroup;
                        else
                        {
                            salesperson = reader.GetString(reader.GetOrdinal("salesperson"));

                            if (!String.IsNullOrEmpty(salespersonGroup))
                                salespersonGroup += ", ";

                            salespersonGroup += salesperson;

                            //  CreateAutoAttachements()
                        }


                    }
                }

            }




            return true;

        }

        private Attachment CreateAutoAttachments(AutoAttachmentTypes autoAttachmentType, Microsoft.Office.Interop.Excel.Application excel, string sprocName, CommissionRecord commissionRecord, string salesperson, Int64 salespersonGroup, Int64 sessionId)
        {
            string separator = "'============";
            string initialValue = "~Initial~";

           
            decimal commissionGroupDescriptionTotal = 0;
            bool hasDataMiningMenuMania = false;
            bool hasDataMiningNewBusiness = false;
            bool hasDataMiningProducts = false;
            bool hasPlaybook = false;
            Int32 rowCounter = 0;
            string attachmentDescription = "";
            string commissionsGroupDescription = initialValue;
            string fileNamePrefix = "";

            excel.Application.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Workbook activeWorkBook = excel.Application.ActiveWorkbook;
            excel.DisplayAlerts = false;

            //remove all worksheets except the first one
            //why are we calling this again? we just called this in the calling method
            while (activeWorkBook.Worksheets.Count > 1)
            {
                Microsoft.Office.Interop.Excel.Worksheet worksheetToDelete = activeWorkBook.Sheets[2];
                worksheetToDelete.Delete();
            }

            excel.DisplayAlerts = true;
            activeWorkBook.Sheets.Add(null, activeWorkBook.Sheets[activeWorkBook.Sheets.Count], null, null);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = activeWorkBook.Sheets[activeWorkBook.Sheets.Count];
            worksheet.Select();

            List<int> rowHeights = new List<int>();

            switch (autoAttachmentType)
            {
                case AutoAttachmentTypes.MenuMania:
                    hasDataMiningMenuMania = true;
                    attachmentDescription = "Data Mining Menu Mania";
                    fileNamePrefix = "Data_Mining_Menu_Mania";

                    rowCounter = 1;

                    //build first header row
                    worksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year;
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    //build second header row
                    worksheet.Cells[rowCounter, 1] = "For " + salesperson + " (" + salespersonGroup + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    //build a third header row
                    worksheet.Cells[rowCounter, 1] = "Data Mining Menu Mania";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    Microsoft.Office.Interop.Excel.Range row = worksheet.Rows[rowCounter];

                    rowHeights.Add(row.RowHeight * 2);

                    Microsoft.Office.Interop.Excel.Range range = worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter];
                    range.MergeCells = true;
                    range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = 1;  //continuous

                    rowCounter++;

                    //add column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 2] = "Commissions Data Mining";
                    FormatCells(worksheet.Columns[2], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 3] = "Amount";
                    FormatCells(worksheet.Columns[3], "$#,##0.00;($#,##0.00)", ExcelHorizontalAlignment.Right, "Currency", false, false, false, false);

                    worksheet.Cells[rowCounter, 4] = "Tran Date";
                    FormatCells(worksheet.Columns[4], "mm/dd/yyyy", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 5] = "Account";
                    FormatCells(worksheet.Columns[5], "#0;(#0)", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 6] = "Client Name";
                    FormatCells(worksheet.Columns[6], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 7] = "Ticket";
                    FormatCells(worksheet.Columns[7], "#0", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    rowCounter++;

                    row = worksheet.Rows[rowCounter];
                    row.Font.Bold = true;
                    row.Font.Underline = ExcelUnderLines.SingleUnderline;

                    //get related commission data
                    using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Data_Mining_Menu_Mania_For_Excel",
                                        new Dictionary<string, object>()
                                        {
                                               { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                               { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                               { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                               { "@pvchrSalesperson", salesperson },
                                        }))
                    {
                        if (!reader.HasRows)
                            return null;

                        decimal groupTotalCommissions = 0;

                        while (reader.Read())
                        {
                            rowCounter++;
                            worksheet.Cells[rowCounter, 1] = reader.GetString(reader.GetOrdinal("salesperson"));
                            worksheet.Cells[rowCounter, 2] = reader.GetString(reader.GetOrdinal("product_commissions_menu_mania_description"));
                            worksheet.Cells[rowCounter, 3] = reader.GetDecimal(reader.GetOrdinal("amount_pretax"));
                            worksheet.Cells[rowCounter, 4] = reader.GetDateTime(reader.GetOrdinal("trandate"));
                            worksheet.Cells[rowCounter, 5] = reader.GetInt32(reader.GetOrdinal("history_core_account"));
                            worksheet.Cells[rowCounter, 6] = reader.GetString(reader.GetOrdinal("clientsdata_clientname"));
                            worksheet.Cells[rowCounter, 7] = reader.GetString(reader.GetOrdinal("history_core_ticket"));
                            groupTotalCommissions += reader.GetDecimal(reader.GetOrdinal("amount_pretax"));
                        }

                        rowCounter++;

                        worksheet.Cells[rowCounter, 3] = separator;

                        rowCounter++;

                        worksheet.Cells[rowCounter, 1] = "TOTALS FOR";
                        worksheet.Cells[rowCounter, 2] = worksheet.Cells[rowCounter - 2, 2].Value;
                        worksheet.Cells[rowCounter, 3] = groupTotalCommissions;
                    }

                    break;
                case AutoAttachmentTypes.NewBusiness:
                    hasDataMiningNewBusiness = true;
                    attachmentDescription = "Data Mining New Business";
                    fileNamePrefix = "Data_Mining_New_Business";

                    rowCounter = 1;

                    worksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year;
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "For " + salesperson + " (" + salespersonGroup + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "Data Mining New Business";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add(row.RowHeight * 2);

                    range = worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter];
                    range.MergeCells = true;
                    range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = 1;  //continuous

                    rowCounter++;

                    //build column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 2] = "Commissions Data Mining";
                    FormatCells(worksheet.Columns[2], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 3] = "Amount";
                    FormatCells(worksheet.Columns[3], "$#,##0.00;($#,##0.00)", ExcelHorizontalAlignment.Right, "Currency", false, false, false, false);

                    worksheet.Cells[rowCounter, 4] = "Tran Date";
                    FormatCells(worksheet.Columns[4], "mm/dd/yyyy", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 5] = "New Business Expiration Date"; ;
                    FormatCells(worksheet.Columns[5], "mm/dd/yyyy", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 6] = "Account";
                    FormatCells(worksheet.Columns[6], "#0;(#0)", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 7] = "Client Name";
                    FormatCells(worksheet.Columns[7], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 8] = "Ticket";
                    FormatCells(worksheet.Columns[8], "#0", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    range = worksheet.Rows[rowCounter];
                    range.Font.Bold = true;
                    range.Font.Underline = ExcelUnderLines.SingleUnderline;

                    //get related commission data
                    using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Data_Mining_New_Business_For_Excel",
                                        new Dictionary<string, object>()
                                        {
                                               { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                               { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                               { "@psdatCommissionsEndDate", commissionRecord.EndDate },
                                               { "@pvchrSalesperson", salesperson },
                                        }))
                    {
                        if (!reader.HasRows)
                            return null;

                        decimal groupTotalCommissions = 0;

                        while (reader.Read())
                        {
                            rowCounter++;
                            worksheet.Cells[rowCounter, 1] = reader.GetString(reader.GetOrdinal("salesperson"));
                            worksheet.Cells[rowCounter, 2] = reader.GetString(reader.GetOrdinal("product_commissions_new_business_description"));
                            worksheet.Cells[rowCounter, 3] = reader.GetDecimal(reader.GetOrdinal("amount_pretax"));
                            worksheet.Cells[rowCounter, 4] = reader.GetDateTime(reader.GetOrdinal("trandate"));
                            worksheet.Cells[rowCounter, 4] = reader.GetString(reader.GetOrdinal("tblcustomfieldsvalues_new_bus_date"));
                            worksheet.Cells[rowCounter, 5] = reader.GetInt32(reader.GetOrdinal("history_core_account"));
                            worksheet.Cells[rowCounter, 6] = reader.GetString(reader.GetOrdinal("clientsdata_clientname"));
                            worksheet.Cells[rowCounter, 7] = reader.GetString(reader.GetOrdinal("history_core_ticket"));
                            groupTotalCommissions += reader.GetDecimal(reader.GetOrdinal("amount_pretax"));
                        }
                    }

                    rowCounter++;

                    worksheet.Cells[rowCounter, 3] = separator;

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "TOTALS FOR";
                    worksheet.Cells[rowCounter, 2] = worksheet.Cells[rowCounter - 2, 2];
                    worksheet.Cells[rowCounter, 3] = commissionGroupDescriptionTotal;

                    break;

                case AutoAttachmentTypes.Playbook:
                    hasPlaybook = true;
                    attachmentDescription = "Playbook";
                    fileNamePrefix = "Playbook";

                    worksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year;
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "For " + salesperson + " (" + salespersonGroup + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "Playbook";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, true);

                    rowCounter += 2;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add(row.RowHeight * 2);

                    List<BarcForExcelRecord> barcForExcelRecords = new List<BarcForExcelRecord>();


                    //possible options: Proc_Select_BARC_Retail_For_Excel, Proc_Select_BARC_Outside_Real_Estate_For_Excel,Proc_Select_BARC_Outside_Auto_For_Excel 
                    using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo." + sprocName,
                    new Dictionary<string, object>()
                    {
                                               { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                               { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                               { "@psdatCommissionsEndDate", commissionRecord.EndDate},
                                               { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")},
                                               { "@pvchrSalesperson", salesperson },
                    }))
                    {
                        if (!reader.HasRows)
                            return null;


                        while (reader.Read())
                        {
                            BarcForExcelRecord barcForExcelRecord = new BarcForExcelRecord();

                            barcForExcelRecord.MarkerFlagName = reader.GetString(reader.GetOrdinal("marker_flag_name"));
                            barcForExcelRecord.GroupDescription = reader.GetString(reader.GetOrdinal("playbook_commissions_groups_description"));
                            barcForExcelRecord.PrintDivisionDescription = reader.GetString(reader.GetOrdinal("playbook_print_division_description"));
                            barcForExcelRecord.RevenueWithoutTaxes = reader.GetDecimal(reader.GetOrdinal("revenue_without_taxes"));
                            barcForExcelRecord.TranDate = reader.GetDateTime(reader.GetOrdinal("trandate"));
                            barcForExcelRecord.Account = reader.GetInt32(reader.GetOrdinal("account"));
                            barcForExcelRecord.ClientName = reader.GetString(reader.GetOrdinal("clientname"));
                            barcForExcelRecord.Pub = reader.GetString(reader.GetOrdinal("pub"));
                            barcForExcelRecord.TranCode = reader.GetString(reader.GetOrdinal("trancode"));
                            barcForExcelRecord.TranType = reader.GetString(reader.GetOrdinal("trantype"));
                            barcForExcelRecord.Ticket = reader.GetInt32(reader.GetOrdinal("ticket"));
                            barcForExcelRecord.SelectSource = reader.GetString(reader.GetOrdinal("select_source"));

                            barcForExcelRecords.Add(barcForExcelRecord);
                        }
                    }

                    using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Marker_Flags",
                                        new Dictionary<string, object>()
                                        {
                                               { "@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName") },
                                               { "@pvchrBuffNewsForBWDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName") },
                                               { "@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName") },
                                               { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")},
                                               { "@pvchrMarkerFlagName", barcForExcelRecords[0].MarkerFlagName }, //this value can't be null, it's hardcoded in the sproc
                                        }))
                    {
                        if (!reader.HasRows)
                            return null;


                        while (reader.Read())
                        {
                            worksheet.Cells[rowCounter, 1] = "Criteria: " + reader.GetString(reader.GetOrdinal("description"));
                        }
                    }

                    range = worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter];
                    range.MergeCells = true;
                    range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = 1;  //continuous

                    rowCounter++;

                    //build column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 2] = "Commissions Playbook";
                    FormatCells(worksheet.Columns[2], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 3] = "Playbook Division";
                    FormatCells(worksheet.Columns[3], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 4] = "Amount";
                    FormatCells(worksheet.Columns[4], "$#,##0.00;($#,##0.00)", ExcelHorizontalAlignment.Right, "Currency", false, false, false, false);

                    worksheet.Cells[rowCounter, 5] = "Tran Date";
                    FormatCells(worksheet.Columns[5], "mm/dd/yyyy", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 6] = "Account";
                    FormatCells(worksheet.Columns[6], "#0;(#0)", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 7] = "Client Name";
                    FormatCells(worksheet.Columns[7], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 8] = "Pub";
                    FormatCells(worksheet.Columns[8], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 9] = "Tran Code";
                    FormatCells(worksheet.Columns[9], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 10] = "Tran Type";
                    FormatCells(worksheet.Columns[10], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 11] = "Ticket";
                    FormatCells(worksheet.Columns[11], "#0", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 12] = "Source";
                    FormatCells(worksheet.Columns[12], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    range = worksheet.Rows[rowCounter];
                    range.Font.Bold = true;
                    range.Font.Underline = ExcelUnderLines.SingleUnderline;

                    //iterate records
                    string commissionGroup = initialValue;
                    foreach (BarcForExcelRecord barcForExcelRecord in barcForExcelRecords)
                    {
                        //add a totals record if we are starting a new group
                        if (barcForExcelRecord.GroupDescription != commissionGroup)
                        {
                            //only add the records if this is not the first pass
                            if (commissionGroup != initialValue)
                            {
                                rowCounter++;

                                worksheet.Cells[rowCounter, 4] = separator;

                                rowCounter++;

                                worksheet.Cells[rowCounter, 1] = "TOTALS FOR";
                                worksheet.Cells[rowCounter, 2] = commissionGroup;
                                worksheet.Cells[rowCounter, 4] = commissionGroupDescriptionTotal;

                                rowCounter += 2;
                            }

                            commissionGroupDescriptionTotal = 0;
                            commissionGroup = barcForExcelRecord.GroupDescription;
                        }


                        //add record
                        rowCounter++;
                        worksheet.Cells[rowCounter, 1] = barcForExcelRecord.Salesperson;
                        worksheet.Cells[rowCounter, 2] = barcForExcelRecord.GroupDescription;
                        worksheet.Cells[rowCounter, 3] = barcForExcelRecord.PrintDivisionDescription;
                        worksheet.Cells[rowCounter, 4] = barcForExcelRecord.RevenueWithoutTaxes;
                        worksheet.Cells[rowCounter, 5] = barcForExcelRecord.TranDate;
                        worksheet.Cells[rowCounter, 6] = barcForExcelRecord.Account;
                        worksheet.Cells[rowCounter, 7] = barcForExcelRecord.ClientName;
                        worksheet.Cells[rowCounter, 8] = barcForExcelRecord.Pub;
                        worksheet.Cells[rowCounter, 9] = barcForExcelRecord.TranCode;
                        worksheet.Cells[rowCounter, 10] = barcForExcelRecord.TranType;
                        worksheet.Cells[rowCounter, 11] = barcForExcelRecord.Ticket;
                        worksheet.Cells[rowCounter, 12] = barcForExcelRecord.SelectSource;
                        commissionGroupDescriptionTotal += barcForExcelRecord.RevenueWithoutTaxes;

                    }

                    //add final record
                    if (commissionsGroupDescription != initialValue)
                    {
                        rowCounter++;

                        worksheet.Cells[rowCounter, 4] = separator;

                        rowCounter++;

                        worksheet.Cells[rowCounter, 1] = "TOTALS FOR";
                        worksheet.Cells[rowCounter, 2] = commissionGroup;
                        worksheet.Cells[rowCounter, 4] = commissionGroupDescriptionTotal;
                    }

                    break;
                case AutoAttachmentTypes.Products:
                    hasDataMiningProducts = true;
                    attachmentDescription = "Data Mining Products";
                    fileNamePrefix = "Data_Mining_Products";

                    //get data
                    List<DataMiningProductForExcel> dataMiningProductForExcels = new List<DataMiningProductForExcel>();

                    using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Data_Mining_Product_For_Excel",
                    new Dictionary<string, object>()
                    {
                                               { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                               { "@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate },
                                               { "@psdatCommissionsEndDate", commissionRecord.EndDate},
                                               { "@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")},
                                               { "@pvchrSalesperson", salesperson },
                    }))
                    {
                        if (!reader.HasRows)
                            return null;


                        while (reader.Read())
                        {
                            DataMiningProductForExcel dataMiningProductForExcelRecord = new DataMiningProductForExcel();

                            dataMiningProductForExcelRecord.Salesperson = reader.GetString(reader.GetOrdinal("salesperson"));
                            dataMiningProductForExcelRecord.GroupDescription = reader.GetString(reader.GetOrdinal("product_commissions_groups_description"));
                            dataMiningProductForExcelRecord.EDNNumber = reader.GetString(reader.GetOrdinal("tbleditions_ednnumber"));
                            dataMiningProductForExcelRecord.Description = reader.GetString(reader.GetOrdinal("tbleditions_descript"));
                            dataMiningProductForExcelRecord.TranDate = reader.GetDateTime(reader.GetOrdinal("trandate"));
                            dataMiningProductForExcelRecord.AmountPreTax = reader.GetDecimal(reader.GetOrdinal("amount_pretax"));
                            dataMiningProductForExcelRecord.HistoryCoreAccount = reader.GetInt32(reader.GetOrdinal("history_core_account"));
                            dataMiningProductForExcelRecord.ClientName = reader.GetString(reader.GetOrdinal("clientsdata_clientname"));
                            dataMiningProductForExcelRecord.HistoryCoreTicket = reader.GetInt32(reader.GetOrdinal("history_core_ticket"));

                            commissionGroupDescriptionTotal = reader.GetDecimal(reader.GetOrdinal("amount_pretax"));

                            dataMiningProductForExcels.Add(dataMiningProductForExcelRecord);
                        }
                    }


                    worksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year;
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "For " + salesperson + " (" + salespersonGroup + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, false);

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "Data Mining Products";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, true);

                    rowCounter += 2;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add(row.RowHeight * 2);

                    FormatCells(worksheet.Cells[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], null, ExcelHorizontalAlignment.Center, null, true, true, false, true);

                    //get descriptions
                    List<string> descriptions = new List<string>();
                    using (SqlDataReader reader = ExecuteQuery(DatabaseConnectionStringNames.Commissions, CommandType.StoredProcedure, "dbo.Proc_Select_Snapshots_Product_Data_Mining_Descriptions",
                                        new Dictionary<string, object>()
                                        {
                                               { "@pintSnapshotsID", commissionRecord.SnapshotId },
                                               { "@pvchrSalesperson", salesperson }
                                        }))
                    {

                        while (reader.Read())
                        {
                            descriptions.Add(reader.GetString(reader.GetOrdinal("tbleditions_ednnumber")));
                        }
                    }


                    worksheet.Cells[rowCounter, 1] = "Selected Data Mining Editions: " + String.Join(", ", descriptions);

                    rowCounter++;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add(row.RowHeight * 2);

                    range = worksheet.Rows[rowCounter];
                    range.MergeCells = true;
                    range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = 1;  //continuous

                    rowCounter++;

                    //build column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 2] = "Commissions Data Mining";
                    FormatCells(worksheet.Columns[2], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 3] = "Data Mining Edition";
                    FormatCells(worksheet.Columns[3], "@", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 4] = "Data Mining Description";
                    FormatCells(worksheet.Columns[4], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 5] = "Amount";
                    FormatCells(worksheet.Columns[5], "$#,##0.00;($#,##0.00)", ExcelHorizontalAlignment.Right, "Currency", false, false, false, false);

                    worksheet.Cells[rowCounter, 6] = "Tran Date";
                    FormatCells(worksheet.Columns[6], "mm/dd/yyyy", ExcelHorizontalAlignment.Center, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 7] = "Account";
                    FormatCells(worksheet.Columns[7], "#0;(#0)", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 8] = "Client Name";
                    FormatCells(worksheet.Columns[8], "@", ExcelHorizontalAlignment.Left, null, false, false, false, false);

                    worksheet.Cells[rowCounter, 9] = "Ticket";
                    FormatCells(worksheet.Columns[9], "#0", ExcelHorizontalAlignment.Left, null, false, false, false, false);


                    range = worksheet.Rows[rowCounter];
                    range.Font.Bold = true;
                    range.Font.Underline = ExcelUnderLines.SingleUnderline;

                    //iterate records
                    string editionDescription = "";
                    string editionNumber = initialValue;
                    commissionGroup = initialValue;
                    foreach (DataMiningProductForExcel dataMiningProductForExcel in dataMiningProductForExcels)
                    {
                        //add a totals record if we are starting a new group
                        if (dataMiningProductForExcel.GroupDescription != commissionGroup)
                        {
                            //only add the records if this is not the first pass
                            if (commissionGroup != initialValue)
                            {
                                rowCounter++;

                                worksheet.Cells[rowCounter, 5] = separator;

                                rowCounter++;

                                worksheet.Cells[rowCounter, 1] = "TOTALS FOR";
                                worksheet.Cells[rowCounter, 2] = commissionGroup;
                                worksheet.Cells[rowCounter, 5] = commissionGroupDescriptionTotal;

                                rowCounter += 2;
                            }

                            commissionGroupDescriptionTotal = 0;
                            commissionGroup = dataMiningProductForExcel.GroupDescription;
                            editionNumber = dataMiningProductForExcel.EDNNumber;
                            editionDescription = dataMiningProductForExcel.Description;
                        }
                        else if (dataMiningProductForExcel.EDNNumber != editionNumber)
                        {
                            if (editionNumber != initialValue)
                                rowCounter++;

                            editionDescription = dataMiningProductForExcel.Description;
                            editionNumber = dataMiningProductForExcel.EDNNumber;
                        }


                        //add record
                        rowCounter++;
                        worksheet.Cells[rowCounter, 1] = dataMiningProductForExcel.Salesperson;
                        worksheet.Cells[rowCounter, 2] = dataMiningProductForExcel.GroupDescription;
                        worksheet.Cells[rowCounter, 3] = dataMiningProductForExcel.EDNNumber;
                        worksheet.Cells[rowCounter, 4] = dataMiningProductForExcel.Description;
                        worksheet.Cells[rowCounter, 5] = dataMiningProductForExcel.AmountPreTax;
                        worksheet.Cells[rowCounter, 6] = dataMiningProductForExcel.TranDate;
                        worksheet.Cells[rowCounter, 7] = dataMiningProductForExcel.HistoryCoreAccount;
                        worksheet.Cells[rowCounter, 8] = dataMiningProductForExcel.ClientName;
                        worksheet.Cells[rowCounter, 9] = dataMiningProductForExcel.HistoryCoreTicket;

                        commissionGroupDescriptionTotal += dataMiningProductForExcel.AmountPreTax;

                    }

                    //add final record
                    if (commissionsGroupDescription != initialValue)
                    {
                        rowCounter++;

                        worksheet.Cells[rowCounter, 5] = separator;

                        rowCounter++;

                        worksheet.Cells[rowCounter, 1] = "TOTALS FOR";
                        worksheet.Cells[rowCounter, 2] = commissionGroup;
                        worksheet.Cells[rowCounter, 5] = commissionGroupDescriptionTotal;
                    }


                    break;
            }


            //set final properties
            worksheet.Columns.AutoFit();
            worksheet.Rows.AutoFit();

            //todo: is this needed?
            //if (rowHeights != null && rowHeights.Count() > 0)
            //{
            //    foreach (int rowHeight in rowHeights)
            //    {
            //        if ()
            //    }
            //}

            excel.PrintCommunication = false;

            worksheet.PageSetup.PrintTitleRows = "$1:$" + rowCounter;
            worksheet.PageSetup.PrintTitleColumns = "";
            worksheet.PageSetup.LeftHeader = "";
            worksheet.PageSetup.CenterHeader = "";
            worksheet.PageSetup.RightHeader = "";
            worksheet.PageSetup.LeftFooter = "";
            worksheet.PageSetup.CenterFooter = "Page &P &N";
            worksheet.PageSetup.RightFooter = "";
            worksheet.PageSetup.LeftMargin = 36;
            worksheet.PageSetup.RightMargin = 36;
            worksheet.PageSetup.TopMargin = 36;
            worksheet.PageSetup.BottomMargin = 36;
            worksheet.PageSetup.HeaderMargin = 18;
            worksheet.PageSetup.FooterMargin = 18;
            worksheet.PageSetup.PrintHeadings = false;
            worksheet.PageSetup.PrintGridlines = false;
            worksheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            worksheet.PageSetup.Zoom = false;
            worksheet.PageSetup.FitToPagesWide = 1;
            worksheet.PageSetup.FitToPagesTall = 999;

            excel.PrintCommunication = true;

            string outputPath = GetConfigurationKeyValue("AttachmentDirectory") + sessionId + "_" + fileNamePrefix + "_" + salesperson + "_" + DateTime.Now.ToString("yyyyMMddhhmmssfff") + ".pdf";

            activeWorkBook.ExportAsFixedFormat(Type: 0, Filename: outputPath);

            return new Attachment()
            {
                Description = attachmentDescription + " For " + salesperson,
                FileName = outputPath,
                HasManiaFlag = hasDataMiningMenuMania,
                HasNewBusinessFlag = hasDataMiningNewBusiness,
                HasProductsFlag = hasDataMiningProducts,
                FileNameExtension = ".pdf",
                FileNamePrefix = fileNamePrefix,
                PlaybookFlag = hasPlaybook,
                Salesperson = salesperson,
                SalespersonGroupId = salespersonGroup
            };

        }

        private void FormatCells(Microsoft.Office.Interop.Excel.Range range, string numberformat, ExcelHorizontalAlignment horizontalAlignment, string style,
                               bool mergeCells, bool isBold, bool isUnderline, bool wrapText)
        {
            if (style != null)
                range.Style = style;

            if (numberformat != null)
                range.NumberFormat = numberformat;

            range.MergeCells = mergeCells;
            range.Font.Bold = isBold;
            range.Font.Underline = isUnderline;
            range.HorizontalAlignment = horizontalAlignment;
        }

        private string ConvertToColumn(Int32 columnNumber)
        {
            Int32 offset = 64;

            if (columnNumber > 256)
                return "";
            else if (columnNumber < 27)
                return ((char)(columnNumber + offset)).ToString();
            else if (columnNumber < 53)
                return "A" + ((char)((columnNumber - 26) + offset)).ToString();
            else if (columnNumber < 79)
                return "B" + ((char)((columnNumber - 52) + offset)).ToString();
            else if (columnNumber < 105)
                return "C" + ((char)((columnNumber - 78) + offset)).ToString();
            else if (columnNumber < 131)
                return "D" + ((char)((columnNumber - 104) + offset)).ToString();
            else if (columnNumber < 157)
                return "E" + ((char)((columnNumber - 130) + offset)).ToString();
            else if (columnNumber < 183)
                return "F" + ((char)((columnNumber - 156) + offset)).ToString();
            else if (columnNumber < 209)
                return "G" + ((char)((columnNumber - 182) + offset)).ToString();
            else if (columnNumber < 235)
                return "H" + ((char)((columnNumber - 208) + offset)).ToString();
            else
                return "I" + ((char)((columnNumber - 234) + offset)).ToString();
        }

        private List<SalespersonGroup> BuildSalespersonGroup(SqlDataReader reader)
        {
            List<SalespersonGroup> salespersonGroups = new List<SalespersonGroup>();

            while (reader.Read())
            {
                SalespersonGroup salespersonGroup = new SalespersonGroup();

                salespersonGroup.SalespersonGroupsId = reader.GetInt32(reader.GetOrdinal("salespersons_groups_id"));
                salespersonGroup.WorksheetName = reader.GetString(reader.GetOrdinal("worksheet_name"));
                salespersonGroup.SalespersonName = reader.GetString(reader.GetOrdinal("salesperson_name"));
                salespersonGroup.TerritoriesId = reader.GetInt32(reader.GetOrdinal("territories_id"));
                salespersonGroup.BARCForExcelStoredProcedure = reader.GetString(reader.GetOrdinal("territory"));
                salespersonGroup.SalespersonCount = reader.GetInt32(reader.GetOrdinal("salesperson_count"));

                salespersonGroups.Add(salespersonGroup);
            }

            return salespersonGroups;
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

