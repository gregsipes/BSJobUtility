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
                    Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Update_CommissionsCreate_Requested").FirstOrDefault();
                    {
                        if (result == null)
                        {
                            WriteToJobLog(JobLogMessageType.INFO, "No commissions create requests exist");
                            return;
                        }

                        //set the commissions id
                        commissionsId = Int64.Parse(result["commissionscreate_requested_id"].ToString());

                        //build log mesage
                        string message = "Processing commissions create request by " + result["requested_user_name"] + " on " +
                                     String.Format("{0:MM/dd/yyyy hh:mm tt}", DateTime.Parse(result["requested_date_time"].ToString()));

                        //todo: do we need the emailsubset process?

                        WriteToJobLog(JobLogMessageType.INFO, message);

                        int month = (Int32)result["commissions_month"];
                        int year = (Int32)result["commissions_year"];
                        Int64 salespersonGroupId = -1;

                        if ((bool)result["new_commissions_flag"])
                        {
                            //this is a new commissions run
                            createType = CommissionCreateTypes.Create;
                            commissionsId = -1;
                        }
                        else if (String.IsNullOrEmpty(result["salespersons_groups_id"].ToString()))
                        {
                            //this is a recreate for structure request
                            createType = CommissionCreateTypes.RecreateForStructure;
                        }
                        else
                        {
                            //this is a recreate for salesperson request
                            createType = CommissionCreateTypes.RecreateForSalesperson;
                            salespersonGroupId = (Int64)result["salespersons_groups_id"];
                        }

                        //create commissions object
                        commissionRecord = new CommissionRecord() { Month = month, Year = year, CommissionsId = commissionsId };
                        commissionRecord.EndDate = (DateTime)result["commissions_end_date"];
                        commissionRecord.MonthStartDate = (DateTime)result["commissions_month_start_date"];
                        commissionRecord.PriorEndDate = (DateTime)result["commissions_prior_end_date"];
                        commissionRecord.PriorMonthStartDate = (DateTime)result["commissions_prior_month_start_date"];
                        commissionRecord.PriorYearStartDate = (DateTime)result["commissions_prior_ytd_start_date"];
                        commissionRecord.YearStartDate = (DateTime)result["commissions_ytd_start_date"];
                        commissionRecord.GainsLossesTopCount = result["gains_losses_top_count"].ToString();
                        commissionRecord.StructuresId = Int64.Parse(result["structures_id"].ToString());
                        commissionRecord.RequestedUserName = result["requested_user_name"].ToString();
                        commissionRecord.SalespersonName = result["salesperson_name"].ToString();
                        commissionRecord.SalespersonGroupId = String.IsNullOrEmpty(result["salespersons_groups_id"].ToString()) ? -1 : Int32.Parse(result["salespersons_groups_id"].ToString());
                    }

                    //process commission request
                    ProcessCommissions(createType, commissionRecord);

                    //todo: build and send email



                    //delete request
                    ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Delete_CommissionsCreate_Requested", new SqlParameter("@pintCommissionsCreateRequestedID", commissionsId));
                }
                catch (Exception ex)
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
            AppConfigSectionName = "CommissionsCreate";
        }

        private void ProcessCommissions(CommissionCreateTypes createType, CommissionRecord commissionsRecord)
        {
            if (createType == CommissionCreateTypes.Create)
                CreateNewCommission(commissionsRecord); //new commissions create request
            else
                RecreateCommission(createType, commissionsRecord);   //recreate a commissions request


            //todo:
            //DeleteAutoAttachments();
        }

        private void CreateNewCommission(CommissionRecord commissionsRecord)
        {
            WriteToJobLog(JobLogMessageType.INFO, "New commissions for " + commissionsRecord.StructuresId.ToString() + " " + commissionsRecord.Month.ToString() + "/" + commissionsRecord.Year);

            //Inserts a new record in the Commissions table and returns a new commissionId (unique value for this run)
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Commissions",
                                                new SqlParameter("@pintStructuresID", commissionsRecord.StructuresId),
                                                new SqlParameter("@pintCommissionsYear", commissionsRecord.Year),
                                                new SqlParameter("@pintCommissionsMonth", commissionsRecord.Month),
                                                new SqlParameter("@psdatCommissionsYTDStartDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.YearStartDate)),
                                                new SqlParameter("@psdatCommissionsMonthStartDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.MonthStartDate)),
                                                new SqlParameter("@psdatCommissionsEndDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.EndDate)),
                                                new SqlParameter("@psdatCommissionsPriorYTDStartDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.PriorYearStartDate)),
                                                new SqlParameter("@psdatCommissionsPriorMonthStartDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.PriorMonthStartDate)),
                                                new SqlParameter("@psdatCommissionsPriorEndDate", String.Format("{0:MM/dd/yyyy}", commissionsRecord.PriorEndDate)),
                                                new SqlParameter("@pintGainsLossesTopCount", String.Format("{0:MM/dd/yyyy}", commissionsRecord.GainsLossesTopCount)),
                                                new SqlParameter("@pvchrUserName", String.Format("{0:MM/dd/yyyy}", commissionsRecord.RequestedUserName))).FirstOrDefault();

            {
                commissionsRecord.SpreadsheetStyle = Int32.Parse(result["spreadsheet_style"].ToString());
                //commissionsRecord.CommissionsId = reader.GetInt64(reader.GetOrdinal("commissions_id"));
                commissionsRecord.SnapshotId = Int64.Parse(result["snapshots_id"].ToString());
                commissionsRecord.PerformanceForBARCInsertStoredProcedure = result["performance_for_barc_insert_stored_procedure"].ToString();
                commissionsRecord.PlaybookForBARCInsertStoredProcedure = result["playbook_for_barc_insert_stored_procedure"].ToString();
                commissionsRecord.PlaybookForBARCUpdateStoredProcedure = result["playbook_for_barc_update_stored_procedure"].ToString();

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

            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Commissions_Recreate",
                                                        new SqlParameter("@pintStructuresID", commissionsRecord.StructuresId),
                                                        new SqlParameter("@pintCommissionsYear", commissionsRecord.Year),
                                                        new SqlParameter("@pintCommissionsMonth", commissionsRecord.Month)).FirstOrDefault();

            if (ValidateProcedure(result, "Commissions cannot be recreated because other commissions are currently being recreated for this structure"))
                return;

            result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Commissions_Paid_Processing",
                                            new SqlParameter("@pintStructuresID", commissionsRecord.StructuresId),
                                            new SqlParameter("@pintCommissionsYear", commissionsRecord.Year),
                                            new SqlParameter("@pintCommissionsMonth", commissionsRecord.Month)).FirstOrDefault();

            if (ValidateProcedure(result, "Commissions cannot be recreated because they are in the process of being paid by Payroll"))
                return;

            result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Structures",
                                            new SqlParameter("@pintStructuresID", commissionsRecord.StructuresId)).FirstOrDefault();

            if (!(bool)result["verified_flag"])
            {
                WriteToJobLog(JobLogMessageType.WARNING, "Structure (" + commissionsRecord.StructuresId + ") must be verified before salesperson's commissions can be recreated");
                return;
            }

            if (createType == CommissionCreateTypes.RecreateForSalesperson)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Commissions_Statuses_Creating",
                                            new SqlParameter("@pintStructuresID", commissionsRecord.StructuresId),
                                            new SqlParameter("@pintSalespersonsGroupsID", commissionsRecord.SalespersonGroupId),
                                            new SqlParameter("@pvchrSalespersonName", commissionsRecord.SalespersonName),
                                            new SqlParameter("@pvchrStatusBy", commissionsRecord.RequestedUserName));
            }
            else
            {
                result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_For_Commissions_Recreate",
                                                              new SqlParameter("@pintCommissionsID", commissionsRecord.CommissionsId),
                                                              new SqlParameter("@pvchrUserName", commissionsRecord.RequestedUserName)).FirstOrDefault();

                if (!(bool)result["creating_flag"])
                {
                    WriteToJobLog(JobLogMessageType.WARNING, "Recreate not creating");
                    return;
                }

                commissionsRecord.SnapshotId = (Int64)result["snapshots_id"];
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

            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.BuffNewsForBW, "dbo.Proc_Select_Commissions_BuffNews_BARC_Populated").FirstOrDefault();
            {
                if (result == null)
                {
                    WriteToJobLog(JobLogMessageType.WARNING, "No BARC data is available for selection");
                    return false;
                }
                else
                    BARCDatetime = (DateTime)result["end_date_time"];
            }


            if (createType != CommissionCreateTypes.Create)
            {

                //build commission object
                result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Commissions_Related",
                                                    new SqlParameter("@pintCommissionsID", commissionRecord.CommissionsId)).FirstOrDefault();

                commissionRecord.EndDate = (DateTime)result["commissions_end_date"];
                commissionRecord.MonthStartDate = (DateTime)result["commissions_month_start_date"];
                commissionRecord.PriorEndDate = (DateTime)result["commissions_prior_end_date"];
                commissionRecord.PriorMonthStartDate = (DateTime)result["commissions_prior_month_start_date"];
                commissionRecord.PriorYearStartDate = (DateTime)result["commissions_prior_ytd_start_date"];
                commissionRecord.YearStartDate = (DateTime)result["commissions_ytd_start_date"];
                commissionRecord.Month = (Int32)result["commissions_month"];
                commissionRecord.Year = (Int32)result["commissions_year"];

                commissionRecord.GainsLossesTopCount = result["gains_losses_top_count"].ToString();
                commissionRecord.SpreadsheetStyle = (Int32)result["spreadsheet_style"];
                commissionRecord.StructuresId = (Int64)result["structures_id"];
                commissionRecord.PerformanceForBARCInsertStoredProcedure = result["performance_for_barc_insert_stored_procedure"].ToString();
                commissionRecord.PlaybookForBARCInsertStoredProcedure = result["playbook_for_barc_insert_stored_procedure"].ToString();
                commissionRecord.PlaybookForBARCUpdateStoredProcedure = result["playbook_for_barc_update_stored_procedure"].ToString();

            }

            if (createType == CommissionCreateTypes.RecreateForSalesperson)
            {
                //set snapshot id (unique id for the run)
                result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots").FirstOrDefault();
                {
                    commissionRecord.SnapshotId = (Int64)result["snapshots_id"];
                }


                result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Commissions_Recreate",
                                                                    new SqlParameter("@pintStructuresID", commissionRecord.StructuresId),
                                                                    new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                                                    new SqlParameter("@pintCommissionsMonth", commissionRecord.Month),
                                                                    new SqlParameter("@psdatCommissionYTDStartDate", commissionRecord.YearStartDate),
                                                                    new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                                                    new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                                                    new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId),
                                                                    new SqlParameter("@pintNewSnapshotsID", commissionRecord.SnapshotId),
                                                                    new SqlParameter("@pvchrRecreateBy", commissionRecord.RequestedUserName),
                                                                    new SqlParameter("@pvchrRecreateComputerName", "")).FirstOrDefault();

                string message = result["message"].ToString();

                if (!String.IsNullOrEmpty(message))
                {
                    WriteToJobLog(JobLogMessageType.WARNING, message);
                    return false;
                }

                commissionsRecreateId = (Int64)result["commissions_recreate_id"];

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
                List<Dictionary<string, object>> salespersonResult = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Salespersons",
                                                         new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId));

                foreach (Dictionary<string, object> record in salespersonResult)
                {
                    salespersons.Add(record["salesperson"].ToString(), record["salesperson_name"].ToString());
                }
            }
            else
            {
                List<Dictionary<string, object>> salespersonResult = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Salespersons_Recreate",
                                                            new SqlParameter("@pintCommissionsRecreateID", commissionRecord.CommissionsId));
                foreach (Dictionary<string, object> record in salespersonResult)
                {
                    salespersons.Add(record["salesperson"].ToString(), record["salesperson_name"].ToString());
                }
            }

            //get commissions inquiry id
            Int64 commissionsInquiriesId = 0;
            result = ExecuteSQL(DatabaseConnectionStringNames.CommissionsRelated, "dbo.Proc_Insert_Commissions_Inquiries",
                                                                new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                                                new SqlParameter("@pintCommissionsMonth", commissionRecord.Month),
                                                                new SqlParameter("@psdatCommissionsYTDStartDate", commissionRecord.YearStartDate),
                                                                new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                                                new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                                                new SqlParameter("@psdatCommissionsPriorYTDStartDate", commissionRecord.PriorYearStartDate),
                                                                new SqlParameter("@psdatCommissionsPriorMonthStartDate", commissionRecord.PriorMonthStartDate),
                                                                new SqlParameter("@psdatCommissionsPriorEndDate", commissionRecord.PriorEndDate),
                                                                new SqlParameter("@pintGainsLossesTopCount", commissionRecord.GainsLossesTopCount)).FirstOrDefault();
            commissionsInquiriesId = Int64.Parse(result["commissions_inquiries_id"].ToString());

            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Product_Data_Mining_Descriptions",
                                                     new SqlParameter("@pvchrtblEditionsEdnNumber", ""));

            foreach (Dictionary<string, object> record in results)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.CommissionsRelated, "dbo.Proc_Insert_Commissions_Inquiries_Data_Mining",
                                            new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                            new SqlParameter("@pvchrtblEditionsEdnNumber", record["tbleditions_ednnumber"].ToString()));
            }

            ExecuteNonQuery(DatabaseConnectionStringNames.BuffNewsForBW, "dbo.Proc_Insert_Commissions_Inquiries",
                                                    new SqlParameter("@pvchrCommissionsRelatedServerInstance", GetConfigurationKeyValue("CommissionsRelatedServerName")),
                                                    new SqlParameter("@pvchrCommissionsRelatedDatabase", GetConfigurationKeyValue("CommissionsRelatedDatabaseName")),
                                                    new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                                    new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                                    new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId));

            foreach (var salesperson in salespersons)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.CommissionsRelated, "dbo.Proc_Insert_Commissions_Inquiries_Responsible_Salespersons",
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key));

                ExecuteNonQuery(DatabaseConnectionStringNames.BuffNewsForBW, "dbo.Proc_Insert_Commissions_Inquiries_Performance_Salespersons",
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key));
            }

            ExecuteNonQuery(DatabaseConnectionStringNames.BuffNewsForBW, "dbo.Proc_Insert_Commissions_Inquiries_Responsible_Salespersons",
                                        new SqlParameter("@pvchrCommissionsRelatedServerInstance", GetConfigurationKeyValue("CommissionsRelatedServerName")),
                                        new SqlParameter("@pvchrCommissionsRelatedDatabase", GetConfigurationKeyValue("CommissionsRelatedDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId));

            WriteToJobLog(JobLogMessageType.INFO, "Selecting menu mania data mining data from Brainworks");

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Data_Mining",
                                        new SqlParameter("@pvchrBrainworksServerInstance", GetConfigurationKeyValue("BrainworksServerName")),
                                        new SqlParameter("@pvchrBrainworksDatabase", GetConfigurationKeyValue("BrainworksDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrStoredProcedure", "Proc_BuffNews_Select_Commissions_Data_Mining_Menu_Mania"));

            WriteToJobLog(JobLogMessageType.INFO, "Selecting new business data mining data from Brainworks");

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Data_Mining",
                                        new SqlParameter("@pvchrBrainworksServerInstance", GetConfigurationKeyValue("BrainworksServerName")),
                                        new SqlParameter("@pvchrBrainworksDatabase", GetConfigurationKeyValue("BrainworksDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrStoredProcedure", "Proc_BuffNews_Select_Commissions_Data_Mining_New_Business"));

            WriteToJobLog(JobLogMessageType.INFO, "Selecting product data mining data from Brainworks");

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Data_Mining",
                                        new SqlParameter("@pvchrBrainworksServerInstance", GetConfigurationKeyValue("BrainworksServerName")),
                                        new SqlParameter("@pvchrBrainworksDatabase", GetConfigurationKeyValue("BrainworksDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrStoredProcedure", "Proc_BuffNews_Select_Commissions_Data_Mining_Product"));

            WriteToJobLog(JobLogMessageType.INFO, "Selecting playbook data from BARC");
            //this is pulling in a snapshot of the BuffNewsForBW.BuffNews_BARC_Brainworks table depending which sproc is passed in
            //Does not create any new records
            //'Proc_Insert_BARC “BWDB\BW,50884', 'BuffNewsForBW', 'CommissionsCreate', '<Cr#@t0rUs3r>', 2607, 'Proc_Select_Commissions_Outside_Auto_Playbook_Detail'
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_BARC",
                                        new SqlParameter("@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName")),
                                        new SqlParameter("@pvchrBuffNewsForBWDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrStoredProcedure", commissionRecord.PlaybookForBARCInsertStoredProcedure));

            WriteToJobLog(JobLogMessageType.INFO, "Selecting performance data from BARC");
            //Does not create any new records
            //Proc_Insert_BARC “BWDB\BW,50884', 'BuffNewsForBW', 'CommissionsCreate', '<Cr#@t0rUs3r>', 2607, 'Proc_Select_Commissions_Outside_Auto_Performance_Detail'
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_BARC",
                                        new SqlParameter("@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName")),
                                        new SqlParameter("@pvchrBuffNewsForBWDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrStoredProcedure", commissionRecord.PerformanceForBARCInsertStoredProcedure));

            WriteToJobLog(JobLogMessageType.INFO, "Selecting gains/losses data from BARC");
            //Creates 631 new records with new snapshots_id.  HOW DID THE SNAPSHOTS ID GET INTO HERE???????
            //Proc_Insert_BARC “BWDB\BW,50884', 'BuffNewsForBW', 'CommissionsCreate', '<Cr#@t0rUs3r>', 2607, 'Proc_Select_Commissions_Gains_Losses_Detail ‘”
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_BARC",
                                        new SqlParameter("@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName")),
                                        new SqlParameter("@pvchrBuffNewsForBWDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pintCommissionsInquiriesID", commissionsInquiriesId),
                                        new SqlParameter("@pvchrStoredProcedure", "Proc_Select_Commissions_Gains_Losses_Detail"));

            WriteToJobLog(JobLogMessageType.INFO, "Initializing snapshots");
            RunSnapshotSprocs(commissionRecord, createType, commissionsRecreateId, commissionRecord.SnapshotId, salespersons);

            return CreateCommissionsSpeadsheets(createType, commissionRecord);

        }

        private bool CreateCommissionsSpeadsheets(CommissionCreateTypes createTypes, CommissionRecord commissionRecord)
        {
            //insert session
            Int64 sessionId = 0;
            Dictionary<string, object> sessionResult = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Sessions",
                                                    new SqlParameter("@pvchrUserName", commissionRecord.RequestedUserName),
                                                    new SqlParameter("@pvchrComputerName", "")).FirstOrDefault();
            sessionId = Int64.Parse(sessionResult["sessions_id"].ToString());

            //build salesperson groups
            List<SalespersonGroup> salespersonGroups = new List<SalespersonGroup>();

            if (createTypes == CommissionCreateTypes.Create)
            {
                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Salespersons_Groups",
                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                        new SqlParameter("@plngTerritoriesID", -1));

                salespersonGroups = BuildSalespersonGroup(results);
            }
            else
            {
                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Salespersons_Groups_Recreate",
                        new SqlParameter("@pintCommissionsRecreateID", commissionRecord.CommissionsId));

                salespersonGroups = BuildSalespersonGroup(results);

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

                List<Dictionary<string, object>> salespersonsResults = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Salespersons",
                                            new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                            new SqlParameter("@pintSalespersonGroupsID", salespersonGroup.SalespersonGroupsId));

                bool isSummaryRecord = false;
                Int64 rowCounter = 0;
                Int64 rowFirstForGroupTotal = 0;
                Int64 rowLastForGroupTotal = 0;
                string currentMonthCommissionsFormula = "";

                foreach (Dictionary<string, object> salespersonResult in salespersonsResults) //while (true) //iterate salespersons 
                {
                    string salesperson = "";
                    string salespersonGroupName = "";
                    if (isSummaryRecord)
                        salesperson = "Summary For " + salespersonGroup;
                    else
                    {
                        salesperson = salespersonResult["salesperson"].ToString();

                        if (!String.IsNullOrEmpty(salespersonGroupName))
                            salespersonGroupName += ", ";

                        salespersonGroupName += salesperson;

                        CreateAutoAttachments(AutoAttachmentTypes.MenuMania, excel, "", commissionRecord, salesperson, (Int32)salespersonResult["salespersons_groups_id"], sessionId, salespersonGroup.SalespersonName);
                        CreateAutoAttachments(AutoAttachmentTypes.NewBusiness, excel, "", commissionRecord, salesperson, (Int32)salespersonResult["salespersons_groups_id"], sessionId, salespersonGroup.SalespersonName);
                        CreateAutoAttachments(AutoAttachmentTypes.Products, excel, "", commissionRecord, salesperson, (Int32)salespersonResult["salespersons_groups_id"], sessionId, salespersonGroup.SalespersonName);
                        CreateAutoAttachments(AutoAttachmentTypes.Playbook, excel, salespersonGroup.BARCForExcelStoredProcedure, commissionRecord, salesperson, (Int32)salespersonResult["salespersons_groups_id"], sessionId, salespersonGroup.SalespersonName);
                    }

                    if (rowCounter != 0)
                        excel.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);

                    activeWorksheet = workbook.Sheets[workbook.Sheets.Count];
                    //  activeWorksheet.Select();

                    activeWorksheet.Name = salespersonGroup.WorksheetName + " " + (isSummaryRecord ? "Summary" : salesperson);
                    //todo: do we need the column width array?

                    rowCounter = 1;

                    activeWorksheet.VPageBreaks.Add(activeWorksheet.Range["G1"]);

                    FormatCells(activeWorksheet.Columns[5], new ExcelFormatOption() { FillColor = ExcelColor.Black });

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(6) + rowCounter],
                                            new ExcelFormatOption()
                                            {
                                                FillColor = ExcelColor.LightGray15,
                                                MergeCells = true,
                                                IsBold = true
                                            });
                    activeWorksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + new DateTime(commissionRecord.Month).ToString("MMM", CultureInfo.InvariantCulture) + " " + commissionRecord.Year;

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(3) + rowCounter], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.LightGray15, IsBold = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });

                    activeWorksheet.Cells[rowCounter, 1] = salespersonGroup.SalespersonName + " (" + salesperson + ")";

                    FormatCells(activeWorksheet.Range[ConvertToColumn(4) + rowCounter + ":" + ConvertToColumn(6) + rowCounter], new ExcelFormatOption() { MergeCells = true, NumberFormat = "@", FillColor = ExcelColor.LightGray15, IsBold = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });

                    activeWorksheet.Cells[rowCounter, 4] = "Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");

                    rowCounter++;

                    SetupWorksheet(excel, activeWorksheet, rowCounter);

                    //todo: store row height?

                    rowCounter++;

                    //FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(6) + rowCounter], new ExcelFormatOption() { IsBold = true, FillColor = ExcelColor.Black, TextColor = ExcelColor.White });

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(5) + rowCounter], new ExcelFormatOption() { IsBold = true, FillColor = ExcelColor.Black, TextColor = ExcelColor.White });
                    activeWorksheet.Cells[rowCounter, 1] = "Playbook Commissions";

                    FormatCells(activeWorksheet.Cells[rowCounter, 6], new ExcelFormatOption() { HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter, IsBold = true, FillColor = ExcelColor.Black, TextColor = ExcelColor.White });
                    activeWorksheet.Cells[rowCounter, 6] = "Goal";

                    rowFirstForGroupTotal = 0;
                    rowLastForGroupTotal = 0;

                    if (isSummaryRecord)
                    {
                        //get playbook groups
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Playbook_Groups_For_Salespersons_Groups_ID",
                                                                                    new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                                     new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["playbook_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["playbook_amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());
                        }
                    }
                    else
                    {
                        //get playbook groups 
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Playbook_Groups_For_Salesperson",
                                                                                new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                                new SqlParameter("@pvchrSalesperson", salesperson));
                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["playbook_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "0.000%;-0.000%", StyleName = "Percent", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = Decimal.Parse(result["percentage"].ToString()) / 100;

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["playbook_amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());
                        }
                    }

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(2) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderLeftLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    activeWorksheet.Cells[rowCounter, 1] = "Total Playbook Commissions";

                    FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });

                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });

                    if (rowFirstForGroupTotal == 0)
                    {
                        activeWorksheet.Cells[rowCounter, 3] = 0;
                        activeWorksheet.Cells[rowCounter, 4] = 0;
                    }
                    else
                    {
                        string formula1 = "";
                        string formula2 = "";
                        Int64 loopCounter = rowFirstForGroupTotal;
                        while (loopCounter < rowLastForGroupTotal)
                        {
                            if (String.IsNullOrEmpty(formula1))
                                formula1 = "=";
                            else
                                formula1 += "+";

                            if (String.IsNullOrEmpty(formula2))
                                formula2 = "=";
                            else
                                formula2 += "+";

                            formula1 = formula1 + "ROUND(" + ConvertToColumn(3) + loopCounter + ",2)";
                            formula2 = formula2 + "ROUND(" + ConvertToColumn(4) + loopCounter + ",2)";

                            loopCounter++;
                        }

                        activeWorksheet.Cells[rowCounter, 3] = formula1;
                        activeWorksheet.Cells[rowCounter, 4] = formula2;
                    }

                    currentMonthCommissionsFormula = "=" + ConvertToColumn(4) + rowCounter;

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderLeftLineStyle = 1, FillColor = ExcelColor.Black, TextColor = ExcelColor.White });
                    activeWorksheet.Cells[rowCounter, 1] = "Product/Goal Based Commissions";

                    rowFirstForGroupTotal = 0;
                    rowLastForGroupTotal = 0;

                    //get products
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_For_Salespersons_Groups_ID",
                                    new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                    new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));


                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());

                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Groups_Product_For_Salesperson",
                                                                            new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                            new SqlParameter("@pvchrSalesperson", salesperson));


                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "0.000%;-0.000%", StyleName = "Percent", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = Decimal.Parse(result["percentage"].ToString()) / 100;

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());

                        }
                    }

                    //build new business
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Groups_New_Business_For_Salespersons_Groups_ID",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));


                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());

                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Groups_New_Business_For_Salesperson",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pvchrSalesperson", salesperson));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "0.000%;-0.000%", StyleName = "Percent", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = Decimal.Parse(result["percentage"].ToString()) / 100;

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());
                        }
                    }

                    //build menu mania
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Groups_Menu_Mania_For_Salespersons_Group_ID",
                                                                       new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                       new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());
                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Groups_Menu_Mania_For_Salesperson",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pvchrSalesperson", salesperson));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[2], new ExcelFormatOption() { NumberFormat = "0.000%;-0.000%", StyleName = "Percent", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = Decimal.Parse(result["percentage"].ToString()) / 100;

                            FormatCells(activeWorksheet.Cells[3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());
                        }
                    }

                    //build product groups other
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Groups_Other_For_Salespersons_Group_ID",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));


                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());

                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Groups_Other_For_Salesperson",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pvchrSalesperson", salesperson));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["product_commissions_groups_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "0.000%;-0.000%", StyleName = "Percent", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = Decimal.Parse(result["percentage"].ToString()) / 100;

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = Decimal.Parse(result["amount"].ToString());

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = Decimal.Parse(result["commission_amount"].ToString());
                        }
                    }

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(2) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderLeftLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    activeWorksheet.Cells[rowCounter, 1] = "Total Product/Goal Based Commissions";

                    FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });
                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });

                    if (rowFirstForGroupTotal == 0)
                    {
                        activeWorksheet.Cells[rowCounter, 3] = 0;
                        activeWorksheet.Cells[rowCounter, 4] = 0;
                    }
                    else
                    {
                        string formula1 = "";
                        string formula2 = "";
                        Int64 loopCounter = rowFirstForGroupTotal;
                        while (loopCounter < rowLastForGroupTotal)
                        {
                            if (String.IsNullOrEmpty(formula1))
                                formula1 = "=";
                            else
                                formula1 += "+";

                            if (String.IsNullOrEmpty(formula2))
                                formula2 = "=";
                            else
                                formula2 += "+";

                            formula1 = formula1 + "ROUND(" + ConvertToColumn(3) + loopCounter + ",2)";
                            formula2 = formula2 + "ROUND(" + ConvertToColumn(4) + loopCounter + ",2)";

                            loopCounter++;
                        }

                        activeWorksheet.Cells[rowCounter, 3] = formula1;
                        activeWorksheet.Cells[rowCounter, 4] = formula2;
                    }

                    currentMonthCommissionsFormula = "=" + ConvertToColumn(4) + rowCounter.ToString();

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { FillColor = ExcelColor.Black, TextColor = ExcelColor.White, BorderLeftLineStyle = 1, IsBold = true });
                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });
                    activeWorksheet.Cells[rowCounter, 1] = "Account Based Commissions";

                    rowFirstForGroupTotal = 0;
                    rowLastForGroupTotal = 0;

                    //get accounts
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Accounts_For_Salespersons_Groups_ID",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["account_description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = (decimal)result["amount"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["commission_amount"];
                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Accounts_For_Salesperson",
                                            new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                            new SqlParameter("@pvchrSalesperson", salesperson));


                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["account_description"];

                            FormatCells(activeWorksheet.Cells[2], new ExcelFormatOption() { NumberFormat = "0.000%;-0.000%", StyleName = "Percent", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = (decimal)result["percentage"] / 100;

                            FormatCells(activeWorksheet.Cells[3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = (decimal)result["amount"];

                            FormatCells(activeWorksheet.Cells[4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["commission_amount"];
                        }
                    }

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(2) + rowCounter], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, BorderLeftLineStyle = 1, IsBold = true, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    activeWorksheet.Cells[rowCounter, 1] = "Total Account Based Commissions";

                    FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });
                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });

                    if (rowFirstForGroupTotal == 0)
                    {
                        activeWorksheet.Cells[rowCounter, 3] = 0;
                        activeWorksheet.Cells[rowCounter, 4] = 0;
                    }
                    else
                    {
                        string formula1 = "";
                        string formula2 = "";
                        Int64 loopCounter = rowFirstForGroupTotal;
                        while (loopCounter < rowLastForGroupTotal)
                        {
                            if (String.IsNullOrEmpty(formula1))
                                formula1 = "=";
                            else
                                formula1 += "+";

                            if (String.IsNullOrEmpty(formula2))
                                formula2 = "=";
                            else
                                formula2 += "+";

                            formula1 = formula1 + "ROUND(" + ConvertToColumn(3) + loopCounter + ",2)";
                            formula2 = formula2 + "ROUND(" + ConvertToColumn(4) + loopCounter + ",2)";

                            loopCounter++;
                        }

                        activeWorksheet.Cells[rowCounter, 3] = formula1;
                        activeWorksheet.Cells[rowCounter, 4] = formula2;
                    }

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(2) + rowCounter], new ExcelFormatOption() { BorderTopLineStyle = 1 });

                    currentMonthCommissionsFormula = "=" + ConvertToColumn(4) + rowCounter.ToString();

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(2) + rowCounter], new ExcelFormatOption() { IsBold = true, MergeCells = true, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { IsBold = true, BorderTopLineStyle = 1, BorderBottomLineStyle = 1, BorderLeftLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    activeWorksheet.Cells[rowCounter, 3] = "Total Sales" + "/r/n" + "Commissions";
                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { IsBold = true, BorderTopLineStyle = 1, BorderBottomLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });
                    activeWorksheet.Cells[rowCounter, 4] = currentMonthCommissionsFormula;
                    currentMonthCommissionsFormula = "=" + ConvertToColumn(4) + rowCounter.ToString();
                    FormatCells(activeWorksheet.Cells[rowCounter, 5], new ExcelFormatOption() { IsBold = true, BorderBottomLineStyle = 1, FillColor = ExcelColor.Black });

                    rowCounter++;

                    //todo: row heights?

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderTopLineStyle = 1, BorderBottomLineStyle = 1, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.Black, TextColor = ExcelColor.White, MergeCells = true });
                    activeWorksheet.Cells[rowCounter, 1] = "Misc. Non-Commission Cash Payments";

                    rowFirstForGroupTotal = 0;
                    rowLastForGroupTotal = 0;

                    //get non-commissions
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Noncommissions_For_Salespersons_Groups_ID",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["amount"];
                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Noncommissions_For_Salesperson",
                                                                                new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                                new SqlParameter("@pvchrSalesperson", salesperson));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["amount"];
                        }
                    }

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(3) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderTopLineStyle = 1, BorderBottomLineStyle = 1, BorderLeftLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    activeWorksheet.Cells[rowCounter, 1] = "Total Misc. Non-Commission Cash Payments";
                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { IsBold = true, BorderRightLineStyle = 1, BorderBottomLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });

                    if (rowFirstForGroupTotal == 0)
                        activeWorksheet.Cells[rowCounter, 4] = 0;
                    else
                        activeWorksheet.Cells[rowCounter, 4] = "=SUM(" + ConvertToColumn(4) + rowFirstForGroupTotal + ":" + ConvertToColumn(4) + rowLastForGroupTotal + ")";


                    currentMonthCommissionsFormula = "=" + ConvertToColumn(4) + rowCounter.ToString();

                    rowCounter++;

                    //todo: row heights?

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderTopLineStyle = 1, BorderBottomLineStyle = 1, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.Black, TextColor = ExcelColor.White, MergeCells = true });
                    activeWorksheet.Cells[rowCounter, 1] = "Chargebacks";

                    rowFirstForGroupTotal = 0;
                    rowLastForGroupTotal = 0;


                    //get chargebacks
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Chargebacks_For_Salespersons_Groups_ID",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["description"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = (decimal)result["amount"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["commission_amount"];

                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Noncommissions_For_Salesperson",
                                                                                new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                                new SqlParameter("@pvchrSalesperson", salesperson));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 1] = result["description"];

                            FormatCells(activeWorksheet.Cells[2], new ExcelFormatOption() { NumberFormat = "0.000%;-0.000%", StyleName = "Percent", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = (decimal)result["percentage"] / 100;

                            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 3] = (decimal)result["amount"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["commission_amount"];

                        }
                    }

                    rowCounter++;

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(3) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderTopLineStyle = 1, BorderBottomLineStyle = 1, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    activeWorksheet.Cells[rowCounter, 1] = "Total Chargebacks";
                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { IsBold = true, BorderRightLineStyle = 1, BorderBottomLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency" });

                    if (rowFirstForGroupTotal == 0)
                        activeWorksheet.Cells[rowCounter, 4] = 0;
                    else
                    {
                        string formula1 = "";
                        Int64 loopCounter = rowFirstForGroupTotal;
                        while (loopCounter < rowLastForGroupTotal)
                        {
                            if (String.IsNullOrEmpty(formula1))
                                formula1 = "=";
                            else
                                formula1 += "+";

                            formula1 = formula1 + "ROUND(" + ConvertToColumn(3) + loopCounter + ",2)";

                            loopCounter++;
                        }

                        activeWorksheet.Cells[rowCounter, 4] = formula1;
                    }

                    currentMonthCommissionsFormula = "=" + ConvertToColumn(4) + rowCounter.ToString();

                    rowCounter++;

                    //todo: rowheights

                    rowCounter++;

                    FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { IsBold = true, BorderLeftLineStyle = 1, FillColor = ExcelColor.Black, TextColor = ExcelColor.White });
                    activeWorksheet.Cells[rowCounter, 1] = "Draw Per Day";
                    FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { IsBold = true, FillColor = ExcelColor.Black, TextColor = ExcelColor.White, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                    activeWorksheet.Cells[rowCounter, 2] = "Number Of Days";
                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { IsBold = true, BorderRightLineStyle = 1, FillColor = ExcelColor.Black, TextColor = ExcelColor.White });
                    activeWorksheet.Cells[rowCounter, 1] = "Monthly Draw";

                    //get draws per day
                    if (isSummaryRecord)
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Draw_Per_Days_For_Salespersons_Groups_ID",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });
                            activeWorksheet.Cells[rowCounter, 1] = (decimal)result["draw_per_day_amount"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = result["number_of_working_days"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["commission_amount"];

                            //todo:
                            //if (String.IsNullOrEmpty(strCurrentMonthCommissionFormula))
                            //    strCurrentMonthCommissionFormula = "=";
                            //else
                            //    strCurrentMonthCommissionFormula = strCurrentMonthCommissionFormula + "-"

                            // strCurrentMonthCommissionFormula = strCurrentMonthCommissionFormula & ConvertToColumn(4) & lngRow
                        }
                    }
                    else
                    {
                        List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Draw_Per_Days_For_Salesperson",
                                                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                        new SqlParameter("@pvchrSalesperson", salesperson));

                        foreach (Dictionary<string, object> result in results)
                        {
                            rowCounter++;

                            if (rowFirstForGroupTotal == 0)
                                rowFirstForGroupTotal = rowCounter;

                            rowLastForGroupTotal = rowCounter;

                            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });
                            activeWorksheet.Cells[rowCounter, 1] = (decimal)result["draw_per_day_amount"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { BorderLeftLineStyle = 1, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 2] = result["number_of_working_days"];

                            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", StyleName = "Currency", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                            activeWorksheet.Cells[rowCounter, 4] = (decimal)result["commission_amount"];

                            //todo:
                            //if (String.IsNullOrEmpty(strCurrentMonthCommissionFormula))
                            //    strCurrentMonthCommissionFormula = "=";
                            //else
                            //    strCurrentMonthCommissionFormula = strCurrentMonthCommissionFormula + "-"

                            // strCurrentMonthCommissionFormula = strCurrentMonthCommissionFormula & ConvertToColumn(4) & lngRow

                        }
                    }

                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { BorderTopLineStyle = 1 });

                    if (isSummaryRecord || salespersonGroups.Count() == 1)
                    {
                        Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Salespersons_Carryover_Summary",
                                                                        new SqlParameter("@pintSalespersonGroupsID", commissionRecord.SalespersonGroupId),
                                                                        new SqlParameter("@pintStructuresID", commissionRecord.StructuresId),
                                                                        new SqlParameter("@pintCommissionsYearCurrent", commissionRecord.Year),
                                                                        new SqlParameter("@pintCommissionsMonthCurrent", commissionRecord.Month)).FirstOrDefault();

                        if (result != null && result.Count() > 0)
                        {
                            decimal priorMonthCommissionAmount = (decimal)result["prior_month_commissions_amount"];
                            decimal priorMonthNonCommissionAmount = (decimal)result["prior_month_noncommissions_amount"];

                            if (priorMonthCommissionAmount < 0 | priorMonthNonCommissionAmount < 0)
                            {
                                rowCounter++;

                                //todo: column heights

                                rowCounter++;

                                FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.Black, TextColor = ExcelColor.White, MergeCells = true });
                                activeWorksheet.Cells[rowCounter, 1] = "Carryover From Prior Month";

                                if (priorMonthCommissionAmount < 0)
                                {
                                    rowCounter++;

                                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(3) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderBottomLineStyle = 1, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                                    activeWorksheet.Cells[rowCounter, 1] = "Commissions Carryover From " + (commissionRecord.Month == 1 ? "12/" + (commissionRecord.Year - 1) : (commissionRecord.Month - 1) + "/" + commissionRecord.Year);

                                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { IsBold = true, BorderBottomLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency", NumberFormat = "$#,##0.00;($#,##0.00)" });
                                    activeWorksheet.Cells[rowCounter, 4] = priorMonthCommissionAmount;

                                    if (String.IsNullOrEmpty(currentMonthCommissionsFormula))
                                        currentMonthCommissionsFormula = "=";
                                    else
                                        currentMonthCommissionsFormula += "+";

                                    currentMonthCommissionsFormula += ConvertToColumn(4) + rowCounter;
                                }

                                if (priorMonthNonCommissionAmount < 0)
                                {
                                    rowCounter++;

                                    FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(3) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderBottomLineStyle = 1, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
                                    activeWorksheet.Cells[rowCounter, 1] = "Misc. Noncommissions Carryover From " + (commissionRecord.Month == 1 ? "12/" + (commissionRecord.Year - 1) : (commissionRecord.Month - 1) + "/" + commissionRecord.Year);

                                    FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { IsBold = true, BorderBottomLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency", NumberFormat = "$#,##0.00;($#,##0.00)" });
                                    activeWorksheet.Cells[rowCounter, 4] = priorMonthCommissionAmount;

                                    if (String.IsNullOrEmpty(currentMonthCommissionsFormula))
                                        currentMonthCommissionsFormula = "=";
                                    else
                                        currentMonthCommissionsFormula += "+";

                                    currentMonthCommissionsFormula += ConvertToColumn(4) + rowCounter;
                                }

                            }
                        }

                        rowCounter++;

                        //todo: column widths

                        rowCounter++;

                        FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.Black, TextColor = ExcelColor.White, MergeCells = true });

                        if (isSummaryRecord || (Int32)salespersonGroup.SalespersonCount == 1)
                            activeWorksheet.Cells[rowCounter, 1] = "Total Compensation This Month";
                        else
                            activeWorksheet.Cells[rowCounter, 1] = "Total Compensation This Month (refer to Summary for totals";

                        rowCounter++;

                        FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(4) + rowCounter], new ExcelFormatOption() { IsBold = true, BorderLeftLineStyle = 1, BorderRightLineStyle = 1, BorderBottomLineStyle = 1, FillColor = ExcelColor.LightGray25, MergeCells = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });

                        if (isSummaryRecord || (Int32)salespersonGroup.SalespersonCount == 1)
                            activeWorksheet.Cells[rowCounter, 1] = "=" + "\"\"" + "Commissions (MINUS Chargebacks, Draw Per Day & Carryover) " + "\"\"" + "&IF(" + ConvertToColumn(4) + rowCounter + "<0," + "\"\"" + "Owed" + "\"\"" + "," + "\"\"" + "Paid By Payroll" + "\"\"" + ")";
                        else
                            activeWorksheet.Cells[rowCounter, 1] = "Commissions (MINUS Chargebacks, Draw Per Day & Carryover) Subtotal";

                        FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { BorderBottomLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency", NumberFormat = "$#,##0.00;($#,##0.00)" });
                        activeWorksheet.Cells[rowCounter, 4] = currentMonthCommissionsFormula;

                        rowCounter++;

                        FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(3) + rowCounter], new ExcelFormatOption() { BorderBottomLineStyle = 1, BorderRightLineStyle = 1, BorderLeftLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, MergeCells = true });

                        if (isSummaryRecord || (Int32)salespersonGroup.SalespersonCount == 1)
                            activeWorksheet.Cells[rowCounter, 1] = "=" + "\"\"" + "Misc. Non-Commission Cash Payments " + "\"\"" + "&IF(" + ConvertToColumn(4) + rowCounter + "<0," + "\"\"" + "Owed" + "\"\"" + "," + "\"\"" + "Paid By Payroll" + "\"\"" + ")";
                        else
                            activeWorksheet.Cells[rowCounter, 1] = "Misc. Non-Commission Cash Payments Subtotal";

                        FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { BorderBottomLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency", NumberFormat = "$#,##0.00;($#,##0.00)" });
                        activeWorksheet.Cells[rowCounter, 4] = currentMonthCommissionsFormula;

                        rowCounter++;

                        FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(3) + rowCounter], new ExcelFormatOption() { BorderBottomLineStyle = 1, BorderRightLineStyle = 1, BorderLeftLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, MergeCells = true });

                        if (isSummaryRecord || (Int32)salespersonGroup.SalespersonCount == 1)
                            activeWorksheet.Cells[rowCounter, 1] = "Total Paid By Payroll";
                        else
                            activeWorksheet.Cells[rowCounter, 1] = "Subtotal";

                        FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { BorderBottomLineStyle = 1, BorderRightLineStyle = 1, FillColor = ExcelColor.LightGray25, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency", NumberFormat = "$#,##0.00;($#,##0.00)" });
                        activeWorksheet.Cells[rowCounter, 4] = "=IF(" + ConvertToColumn(4) + rowCounter + ">0," + ConvertToColumn(4) + rowCounter + ",0)+IF(" + ConvertToColumn(4) + rowCounter + ">0," + ConvertToColumn(4) + rowCounter + ",0)";

                        activeWorksheet.Columns[4].Autofit();
                        activeWorksheet.HPageBreaks.Add(activeWorksheet.Cells["A" + (rowCounter + 1)]);

                        if (isSummaryRecord)
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Update_Snapshots_Salespersons_Groups_Amounts",
                                                new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                new SqlParameter("@pintSalespersonsGroupsID", commissionRecord.SalespersonGroupId),
                                                new SqlParameter("@pmnyCurrentMonthCommissionsCarryover", (decimal)(activeWorksheet.Range[ConvertToColumn(4) + rowCounter]).Text()),
                                                new SqlParameter("@pmnyCurrentMonthNoncommissionsCarryover", (decimal)(activeWorksheet.Range[ConvertToColumn(4) + rowCounter]).Text()));
                        }
                        else
                        {
                            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Update_Snapshots_Salespersons_Amounts",
                                                new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                new SqlParameter("@pvchrSalesperson", salesperson),
                                                new SqlParameter("@pmnyCurrentMonthCommissionsCarryover", (decimal)(activeWorksheet.Range[ConvertToColumn(4) + rowCounter]).Text()),
                                                new SqlParameter("@pmnyCurrentMonthNoncommissionsCarryover", (decimal)(activeWorksheet.Range[ConvertToColumn(4) + rowCounter]).Text()));
                        }

                    }

                    int columnCounter = 1;
                    activeWorksheet.Columns.AutoFit();
                    foreach (Microsoft.Office.Interop.Excel.Range column in activeWorksheet.Columns)
                    {
                        if (columnCounter == 1)
                            column.ColumnWidth = 40;
                        else if (columnCounter == 2)
                            column.ColumnWidth = 15;
                        else if (columnCounter == 3)
                            column.ColumnWidth = 15;
                        else
                            column.ColumnWidth = 100;
                    }

                    //todo: do we need this
                    //             If intColumnWidthsIndex<> -1 Then
                    //             For intColumnWidthsIndex = 0 To audfColumnWidths.GetUpperBound(0)
                    //                     objExcelColumn = .Columns(audfColumnWidths(intColumnWidthsIndex).intColumn)
                    //                     objExcelColumn.ColumnWidth = audfColumnWidths(intColumnWidthsIndex).decColumnWidth
                    //                 Next
                    //             End If

                    //             If intRowHeightsIndex <> -1 Then
                    //                 For intRowHeightsIndex = 0 To audfRowHeights.GetUpperBound(0)
                    //                     objExcelRow = .Rows(audfRowHeights(intRowHeightsIndex).lngRow)
                    //                     objExcelRow.RowHeight = audfRowHeights(intRowHeightsIndex).decRowHeight
                    //                     objExcelRow.VerticalAlignment = mcintCenterVertical
                    //                 Next
                    //             End If


                    //build performance summary
                    if (!isSummaryRecord)
                        BuildPerformanceSummary(excel, activeWorksheet, commissionRecord, salespersonGroup.SalespersonGroupsId, salesperson, salespersonGroupName, Decimal.Parse(salespersonResult["performance_goal_percentage"].ToString()));
                
                }
                activeWorksheet = workbook.Sheets[1];
                activeWorksheet.Activate();

                string fileName = GetConfigurationKeyValue("AttachmentDirectory") + sessionId + "_SPG_" + salespersonGroup.SalespersonGroupsId + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".xlsx";
                workbook.SaveAs(FileFormat: 51, Filename: fileName);
                workbook.Close(SaveChanges: false);
            }

            return true;

        }

        private void BuildPerformanceSummary(Microsoft.Office.Interop.Excel.Application excel, Microsoft.Office.Interop.Excel.Worksheet activeWorksheet, CommissionRecord commissionRecord, Int64 salespersonGroupId, string salespersonName, string salesperson, decimal performanceGoalPercentage)
        {
            WriteToJobLog(JobLogMessageType.INFO,  "Creating performance summary attachement for " + salespersonName + " (" + salesperson + ")");

            excel.Application.Workbooks.Add();
            excel.Application.DisplayAlerts = false;

            //todo: do we need to check the sheet count again?

            excel.Application.DisplayAlerts = true;

            WriteToJobLog(JobLogMessageType.INFO, "Started select of BARC performance summary data");

            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_BARC_Performance",
                                                                    new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                                                    new SqlParameter("@psdatCurrentYTDStartDate", commissionRecord.YearStartDate),
                                                                    new SqlParameter("@psdatCurrentMonthStartDate", commissionRecord.MonthStartDate),
                                                                    new SqlParameter("@psdatCurrentEndDate", commissionRecord.EndDate),
                                                                    new SqlParameter("@psdatPriorYTDStartDate", commissionRecord.PriorYearStartDate),
                                                                    new SqlParameter("@psdatPriorMonthStartDate", commissionRecord.PriorMonthStartDate),
                                                                    new SqlParameter("@psdatPriorEndDate", commissionRecord.PriorEndDate),
                                                                    new SqlParameter("@pvchrSalesperson", salesperson)).FirstOrDefault();

            Decimal monthRevenueCurrent = result["month_revenue_current"] == null ? 0 : Decimal.Parse(result["month_revenue_current"].ToString());
            Decimal monthRevenuePrior = result["month_revenue_prior"] == null ? 0 : Decimal.Parse(result["month_revenue_prior"].ToString());
            Decimal ytdRevenueCurrent = result["ytd_revenue_current"] == null ? 0 : Decimal.Parse(result["ytd_revenue_current"].ToString());
            Decimal ytdRevenuePrior = result["ytd_revenue_prior"] == null ? 0 : Decimal.Parse(result["ytd_revenue_prior"].ToString());
            Decimal monthActiveAccountsCurrent = result["month_active_accounts_current"] == null ? 0 : Decimal.Parse(result["month_active_accounts_current"].ToString());
            Decimal monthActiveAccountsPrior = result["month_active_accounts_prior"] == null ? 0 : Decimal.Parse(result["month_active_accounts_prior"].ToString());

            activeWorksheet.Name = "Performance Summary";

            int rowCounter = 1;

            activeWorksheet.Cells[rowCounter, 1] = "TBN Salesperson Performance Summary For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year;
            FormatCells(activeWorksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(5) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true, FillColor = ExcelColor.LightGray15 });

            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.LightGray15, IsBold = true });
            activeWorksheet.Cells[rowCounter, 1] = salespersonName + " (" + salesperson + ")";

            FormatCells(activeWorksheet.Range[ConvertToColumn(2) + rowCounter + ":" + ConvertToColumn(5) + rowCounter], new ExcelFormatOption() { MergeCells = true, NumberFormat = "@", FillColor = ExcelColor.LightGray15, IsBold = true, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 4] = "Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");

            rowCounter++;

            //todo: row heights

            rowCounter++;

            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() {NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White, IsBold = true});
            activeWorksheet.Cells[rowCounter, 1] = "Performance";

            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White, IsBold = true });
            activeWorksheet.Cells[rowCounter, 2] = commissionRecord.Year;

            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White, IsBold = true });
            activeWorksheet.Cells[rowCounter, 3] = commissionRecord.Year - 1;

            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White, IsBold = true });
            activeWorksheet.Cells[rowCounter, 4] = "Variance";

            FormatCells(activeWorksheet.Cells[rowCounter, 5], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White, IsBold = true });
            activeWorksheet.Cells[rowCounter, 5] = "%";

            rowCounter++;

            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 1] = "Monthly Actual";

            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 2] = monthRevenueCurrent;

            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 3] = monthRevenuePrior;

            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 4] = "=" + ConvertToColumn(2) + rowCounter + "-" + ConvertToColumn(3) + rowCounter;

            FormatCells(activeWorksheet.Cells[rowCounter, 5], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 5] = "=if(" + ConvertToColumn(3) + rowCounter + "=0,if(" + ConvertToColumn(2) + rowCounter + "=0,0,if(" + ConvertToColumn(2) + rowCounter + "<0,-1,1))," + ConvertToColumn(4) + rowCounter +
                                                    "/abs(" + ConvertToColumn(3) + rowCounter + "))";

            rowCounter++;

            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 1] = "YTD Actual";

            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 2] = ytdRevenueCurrent;

            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 3] = ytdRevenuePrior;

            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 4] = "=" + ConvertToColumn(2) + rowCounter + "-" + ConvertToColumn(3) + rowCounter;

            FormatCells(activeWorksheet.Cells[rowCounter, 5], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 5] = "=if(" + ConvertToColumn(3) + rowCounter + "=0,if(" + ConvertToColumn(2) + rowCounter + "=0,0,if(" + ConvertToColumn(2) + rowCounter + "<0,-1,1))," + ConvertToColumn(4) + rowCounter +
                                                    "/abs(" + ConvertToColumn(3) + rowCounter + "))";

            rowCounter++;

            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 1] = "Monthly Active Accounts";

            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 2] = monthActiveAccountsCurrent;

            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 3] = monthActiveAccountsPrior;

            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 4] = "=" + ConvertToColumn(2) + rowCounter + "-" + ConvertToColumn(3) + rowCounter;
            
            FormatCells(activeWorksheet.Cells[rowCounter, 5], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 5] = "=if(" + ConvertToColumn(3) + rowCounter + "=0,if(" + ConvertToColumn(2) + rowCounter + "=0,0,if(" + ConvertToColumn(2) + rowCounter + "<0,-1,1))," + ConvertToColumn(4) + rowCounter +
                                                    "/abs(" + ConvertToColumn(3) + rowCounter + "))";

            rowCounter++;

            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White });
            activeWorksheet.Cells[rowCounter, 1] = "Actual vs. Goal";

            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 2] = "Actual";

            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White });
            activeWorksheet.Cells[rowCounter, 3] = "Goal (" + performanceGoalPercentage.ToString("P2");

            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White, HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 4] = "Variance";

            FormatCells(activeWorksheet.Cells[rowCounter, 5], new ExcelFormatOption() { NumberFormat = "@", FillColor = ExcelColor.Black, TextColor = ExcelColor.White});
            activeWorksheet.Cells[rowCounter, 5] = "%";

            rowCounter++;


            FormatCells(activeWorksheet.Cells[rowCounter, 1], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight });
            activeWorksheet.Cells[rowCounter, 1] = commissionRecord.Year + " YTD";
            
            FormatCells(activeWorksheet.Cells[rowCounter, 2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 2] = ytdRevenueCurrent;

            FormatCells(activeWorksheet.Cells[rowCounter, 3], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 3] = ytdRevenuePrior + (ytdRevenuePrior * (performanceGoalPercentage / 100));

            FormatCells(activeWorksheet.Cells[rowCounter, 4], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 4] = "=" + ConvertToColumn(2) + rowCounter + "-" + ConvertToColumn(3) + rowCounter;

            FormatCells(activeWorksheet.Cells[rowCounter, 5], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });
            activeWorksheet.Cells[rowCounter, 5] = "=if(" + ConvertToColumn(3) + rowCounter + "=0,if(" + ConvertToColumn(2) + rowCounter + "=0,0,if(" + ConvertToColumn(2) + rowCounter + "<0,-1,1))," + ConvertToColumn(4) + rowCounter +
                                                    "/abs(" + ConvertToColumn(3) + rowCounter + "))";

            rowCounter++;

            //todo: row heights

            rowCounter++;



        }


        private void SetupWorksheet(Microsoft.Office.Interop.Excel.Application excel, Microsoft.Office.Interop.Excel.Worksheet worksheet, Int64 rowCounter)
        {
            excel.PrintCommunication = false;

            if (rowCounter > 0)
                worksheet.PageSetup.PrintTitleRows = "$1:$" + rowCounter;
            else
                worksheet.PageSetup.PrintTitleRows = "";

            worksheet.PageSetup.PrintTitleColumns = "";
            worksheet.PageSetup.LeftHeader = "";
            worksheet.PageSetup.CenterHeader = "";
            worksheet.PageSetup.RightHeader = "";
            worksheet.PageSetup.LeftFooter = "";
            worksheet.PageSetup.CenterFooter = "";
            worksheet.PageSetup.RightFooter = "";
            worksheet.PageSetup.LeftMargin = 36; //0.5 inches
            worksheet.PageSetup.RightMargin = 36; //0.5 inches
            worksheet.PageSetup.TopMargin = 36; //0.5 inches
            worksheet.PageSetup.BottomMargin = 36; //0.5 inches
            worksheet.PageSetup.HeaderMargin = 18; //0.25 inches
            worksheet.PageSetup.FooterMargin = 18; //0.25 inches
            worksheet.PageSetup.PrintHeadings = false;
            worksheet.PageSetup.PrintGridlines = false;
            worksheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
            worksheet.PageSetup.Zoom = 90;

            excel.PrintCommunication = true;
        }

        private Attachment CreateAutoAttachments(AutoAttachmentTypes autoAttachmentType, Microsoft.Office.Interop.Excel.Application excel, string sprocName, CommissionRecord commissionRecord, string salesperson,
                                                   Int32 salespersonGroupId, Int64 sessionId, string salespersonGroupName)
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
            activeWorkBook.Sheets.Add(After: activeWorkBook.Sheets[activeWorkBook.Sheets.Count]);
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
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });
                    rowCounter++;

                    //build second header row
                    worksheet.Cells[rowCounter, 1] = "For " + salespersonGroupName + " (" + salesperson + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    //build a third header row
                    worksheet.Cells[rowCounter, 1] = "Data Mining Menu Mania";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    Microsoft.Office.Interop.Excel.Range row = worksheet.Rows[rowCounter];

                    rowHeights.Add((Int32)row.RowHeight * 2);

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], new ExcelFormatOption() { MergeCells = true, BorderTopLineStyle = 1 });

                    rowCounter++;

                    //add column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], new ExcelFormatOption() { NumberFormat = "@" });

                    worksheet.Cells[rowCounter, 2] = "Commissions Data Mining";
                    FormatCells(worksheet.Columns[2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 3] = "Amount";
                    FormatCells(worksheet.Columns[3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });

                    worksheet.Cells[rowCounter, 4] = "Tran Date";
                    FormatCells(worksheet.Columns[4], new ExcelFormatOption() { NumberFormat = "mm/dd/yyyy" });

                    worksheet.Cells[rowCounter, 5] = "Account";
                    FormatCells(worksheet.Columns[5], new ExcelFormatOption() { NumberFormat = "#0;(#0)" });

                    worksheet.Cells[rowCounter, 6] = "Client Name";
                    FormatCells(worksheet.Columns[6], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 7] = "Ticket";
                    FormatCells(worksheet.Columns[7], new ExcelFormatOption() { NumberFormat = "#0", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    rowCounter++;

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(7) + rowCounter], new ExcelFormatOption() { IsBold = true, IsUnderLine = true });

                    //get related commission data
                    List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Data_Mining_Menu_Mania_For_Excel",
                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                        new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                        new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                        new SqlParameter("@pvchrSalesperson", salesperson));

                    if (results == null || results.Count <= 0)
                        return null;

                    decimal groupTotalCommissions = 0;

                    foreach (Dictionary<string, object> result in results)
                    {
                        rowCounter++;
                        worksheet.Cells[rowCounter, 1] = result["salesperson"].ToString();
                        worksheet.Cells[rowCounter, 2] = result["product_commissions_menu_mania_description"].ToString();
                        worksheet.Cells[rowCounter, 3] = (decimal)result["amount_pretax"];
                        worksheet.Cells[rowCounter, 4] = (DateTime)result["trandate"];
                        worksheet.Cells[rowCounter, 5] = (Int32)result["history_core_account"];
                        worksheet.Cells[rowCounter, 6] = result["clientsdata_clientname"].ToString();
                        worksheet.Cells[rowCounter, 7] = result["history_core_ticket"].ToString();
                        groupTotalCommissions += (decimal)result["amount_pretax"];
                    }

                    rowCounter++;

                    worksheet.Cells[rowCounter, 3] = separator;

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "TOTALS FOR";
                    worksheet.Cells[rowCounter, 2] = worksheet.Cells[rowCounter - 2, 2].Value;
                    worksheet.Cells[rowCounter, 3] = groupTotalCommissions;

                    break;
                case AutoAttachmentTypes.NewBusiness:
                    hasDataMiningNewBusiness = true;
                    attachmentDescription = "Data Mining New Business";
                    fileNamePrefix = "Data_Mining_New_Business";

                    rowCounter = 1;

                    worksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year;
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "For " + salespersonGroupName + " (" + salesperson + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "Data Mining New Business";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add((Int32)row.RowHeight * 2);

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], new ExcelFormatOption() { MergeCells = true, BorderTopLineStyle = 1 });

                    rowCounter++;

                    //build column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], new ExcelFormatOption() { NumberFormat = "@" });

                    worksheet.Cells[rowCounter, 2] = "Commissions Data Mining";
                    FormatCells(worksheet.Columns[2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 3] = "Amount";
                    FormatCells(worksheet.Columns[3], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });

                    worksheet.Cells[rowCounter, 4] = "Tran Date";
                    FormatCells(worksheet.Columns[4], new ExcelFormatOption() { NumberFormat = "mm/dd/yyyy" });

                    worksheet.Cells[rowCounter, 5] = "New Business Expiration Date";
                    FormatCells(worksheet.Columns[5], new ExcelFormatOption() { NumberFormat = "mm/dd/yyyy" });

                    worksheet.Cells[rowCounter, 6] = "Account";
                    FormatCells(worksheet.Columns[6], new ExcelFormatOption() { NumberFormat = "#0;(#0)", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 7] = "Client Name";
                    FormatCells(worksheet.Columns[7], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 8] = "Ticket";
                    FormatCells(worksheet.Columns[8], new ExcelFormatOption() { NumberFormat = "#0", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(8) + rowCounter], new ExcelFormatOption() { IsBold = true, IsUnderLine = true });

                    //get related commission data
                    results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Data_Mining_New_Business_For_Excel",
                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                        new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                        new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                        new SqlParameter("@pvchrSalesperson", salesperson));

                    if (results != null || results.Count <= 0)
                        return null;

                    groupTotalCommissions = 0;

                    foreach (Dictionary<string, object> result in results)
                    {
                        rowCounter++;
                        worksheet.Cells[rowCounter, 1] = result["salesperson"].ToString();
                        worksheet.Cells[rowCounter, 2] = result["product_commissions_new_business_description"].ToString();
                        worksheet.Cells[rowCounter, 3] = result["amount_pretax"].ToString();
                        worksheet.Cells[rowCounter, 4] = (DateTime)result["trandate"];
                        worksheet.Cells[rowCounter, 4] = result["tblcustomfieldsvalues_new_bus_date"].ToString();
                        worksheet.Cells[rowCounter, 5] = (Int32)result["history_core_account"];
                        worksheet.Cells[rowCounter, 6] = result["clientsdata_clientname"].ToString();
                        worksheet.Cells[rowCounter, 7] = result["history_core_ticket"].ToString();
                        groupTotalCommissions += (decimal)result["amount_pretax"];
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

                    rowCounter = 1;

                    worksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year.ToString();
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "For " + salespersonGroupName + " (" + salesperson + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "Playbook";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter += 2;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add((Int32)row.RowHeight * 2);

                    List<BarcForExcelRecord> barcForExcelRecords = new List<BarcForExcelRecord>();


                    //possible options: Proc_Select_BARC_Retail_For_Excel, Proc_Select_BARC_Outside_Real_Estate_For_Excel,Proc_Select_BARC_Outside_Auto_For_Excel 
                    results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo." + sprocName,
                                               new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                               new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                               new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                               new SqlParameter("@pvchrSalesperson", salesperson));

                    if (results == null || results.Count() <= 0)
                        return null;


                    foreach (Dictionary<string, object> result in results)
                    {
                        BarcForExcelRecord barcForExcelRecord = new BarcForExcelRecord();

                        barcForExcelRecord.MarkerFlagName = result["marker_flag_name"].ToString();
                        barcForExcelRecord.GroupDescription = result["playbook_commissions_groups_description"].ToString();
                        barcForExcelRecord.PrintDivisionDescription = result["playbook_print_division_description"].ToString();
                        barcForExcelRecord.RevenueWithoutTaxes = (decimal)result["revenue_without_taxes"];
                        barcForExcelRecord.TranDate = (DateTime)result["trandate"];
                        barcForExcelRecord.Account = (Int32)result["account"];
                        barcForExcelRecord.ClientName = result["clientname"].ToString();
                        barcForExcelRecord.Pub = result["pub"].ToString();
                        barcForExcelRecord.TranCode = result["trancode"].ToString();
                        barcForExcelRecord.TranType = result["trantype"].ToString();
                        barcForExcelRecord.Ticket = (Int32)result["ticket"];
                        barcForExcelRecord.SelectSource = result["select_source"].ToString();
                        barcForExcelRecord.Salesperson = result["salesperson"].ToString();

                        barcForExcelRecords.Add(barcForExcelRecord);
                    }


                    results = ExecuteSQL(DatabaseConnectionStringNames.BARC, "dbo.Proc_Select_Marker_Flags",
                                        new SqlParameter("@pvchrBuffNewsForBWServerInstance", GetConfigurationKeyValue("BuffNewsForBWServerName")),
                                        new SqlParameter("@pvchrBuffNewsForBWDatabase", GetConfigurationKeyValue("BuffNewsForBWDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("CommissionsRelatedUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("CommissionsRelatedPassword")),
                                        new SqlParameter("@pvchrMarkerFlagName", barcForExcelRecords[0].MarkerFlagName)); //this value can't be null, it's hardcoded in the sproc

                    if (results == null || results.Count() <= 0)
                        return null;

                   FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], new ExcelFormatOption() { MergeCells = true, BorderBottomLineStyle = 1 });

                    foreach (Dictionary<string, object> result in results)
                    {
                        worksheet.Cells[rowCounter, 1] = "Criteria: " + result["description"].ToString();
                    }

                    rowCounter+= 2;

                    //build column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], new ExcelFormatOption() { NumberFormat = "@" });

                    worksheet.Cells[rowCounter, 2] = "Commissions Playbook";
                    FormatCells(worksheet.Columns[2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 3] = "Playbook Division";
                    FormatCells(worksheet.Columns[3], new ExcelFormatOption() { NumberFormat = "@" });

                    worksheet.Cells[rowCounter, 4] = "Amount";
                    FormatCells(worksheet.Columns[4], new ExcelFormatOption() { StyleName = "Currency" });

                    worksheet.Cells[rowCounter, 5] = "Tran Date";
                    FormatCells(worksheet.Columns[5], new ExcelFormatOption() { NumberFormat = "mm/dd/yyyy" });

                    worksheet.Cells[rowCounter, 6] = "Account";
                    FormatCells(worksheet.Columns[6], new ExcelFormatOption() { NumberFormat = "#0;(#0)" });

                    worksheet.Cells[rowCounter, 7] = "Client Name";
                    FormatCells(worksheet.Columns[7], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 8] = "Pub";
                    FormatCells(worksheet.Columns[8], new ExcelFormatOption() { NumberFormat = "@" });

                    worksheet.Cells[rowCounter, 9] = "Tran Code";
                    FormatCells(worksheet.Columns[9], new ExcelFormatOption() { NumberFormat = "@" });

                    worksheet.Cells[rowCounter, 10] = "Tran Type";
                    FormatCells(worksheet.Columns[10], new ExcelFormatOption() { NumberFormat = "@" });

                    worksheet.Cells[rowCounter, 11] = "Ticket";
                    FormatCells(worksheet.Columns[11], new ExcelFormatOption() { NumberFormat = "#0", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 12] = "Source";
                    FormatCells(worksheet.Columns[12], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(12) + rowCounter], new ExcelFormatOption() { IsBold = true, IsUnderLine = true });


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
                    if (commissionGroup != initialValue)
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

                    results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Data_Mining_Product_For_Excel",
                                            new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                            new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                            new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                            new SqlParameter("@pvchrSalesperson", salesperson));

                    if (results == null || results.Count() <= 0)
                        return null;


                    foreach (Dictionary<string, object> result in results)
                    {
                        DataMiningProductForExcel dataMiningProductForExcelRecord = new DataMiningProductForExcel();

                        dataMiningProductForExcelRecord.Salesperson = result["salesperson"].ToString();
                        dataMiningProductForExcelRecord.GroupDescription = result["product_commissions_groups_description"].ToString();
                        dataMiningProductForExcelRecord.EDNNumber = result["tbleditions_ednnumber"].ToString();
                        dataMiningProductForExcelRecord.Description = result["tbleditions_descript"].ToString();
                        dataMiningProductForExcelRecord.TranDate = (DateTime)result["trandate"];
                        dataMiningProductForExcelRecord.AmountPreTax = (decimal)result["amount_pretax"];
                        dataMiningProductForExcelRecord.HistoryCoreAccount = (Int32)result["history_core_account"];
                        dataMiningProductForExcelRecord.ClientName = result["clientsdata_clientname"].ToString();
                        dataMiningProductForExcelRecord.HistoryCoreTicket = (Int32)result["history_core_ticket"];

                        commissionGroupDescriptionTotal = (decimal)result["amount_pretax"];

                        dataMiningProductForExcels.Add(dataMiningProductForExcelRecord);
                    }


                    worksheet.Cells[rowCounter, 1] = "TBN Salesperson Commissions For " + CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(commissionRecord.Month) + " " + commissionRecord.Year;
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "For " + salespersonGroupName + " (" + salesperson + ") Created " + DateTime.Now.ToString("MM/dd/yyyy hh:mm tt");
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter++;

                    worksheet.Cells[rowCounter, 1] = "Data Mining Products";
                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    rowCounter += 2;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add((Int32)row.RowHeight * 2);

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], new ExcelFormatOption() { MergeCells = true, IsBold = true });

                    //get descriptions
                    List<string> descriptions = new List<string>();
                    results = ExecuteSQL(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Select_Snapshots_Product_Data_Mining_Descriptions",
                                        new SqlParameter("@pintSnapshotsID", commissionRecord.SnapshotId),
                                        new SqlParameter("@pvchrSalesperson", salesperson));

                    foreach (Dictionary<string, object> result in results)
                    {
                        descriptions.Add(result["tbleditions_ednnumber"].ToString());
                    }

                    worksheet.Cells[rowCounter, 1] = "Selected Data Mining Editions: " + String.Join(", ", descriptions);

                    rowCounter++;

                    row = worksheet.Rows[rowCounter];
                    rowHeights.Add((Int32)row.RowHeight * 2);

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], new ExcelFormatOption() { MergeCells = true, BorderTopLineStyle = 1 });

                    rowCounter++;

                    //build column headers
                    worksheet.Cells[rowCounter, 1] = "Commissions Salesperson";
                    FormatCells(worksheet.Columns[1], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter });

                    worksheet.Cells[rowCounter, 2] = "Commissions Data Mining";
                    FormatCells(worksheet.Columns[2], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 3] = "Data Mining Edition";
                    FormatCells(worksheet.Columns[3], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter });

                    worksheet.Cells[rowCounter, 4] = "Data Mining Description";
                    FormatCells(worksheet.Columns[4], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 5] = "Amount";
                    FormatCells(worksheet.Columns[5], new ExcelFormatOption() { NumberFormat = "$#,##0.00;($#,##0.00)", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight, StyleName = "Currency" });

                    worksheet.Cells[rowCounter, 6] = "Tran Date";
                    FormatCells(worksheet.Columns[6], new ExcelFormatOption() { NumberFormat = "mm/dd/yyyy", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter });

                    worksheet.Cells[rowCounter, 7] = "Account";
                    FormatCells(worksheet.Columns[7], new ExcelFormatOption() { NumberFormat = "#0;(#0)", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 8] = "Client Name";
                    FormatCells(worksheet.Columns[8], new ExcelFormatOption() { NumberFormat = "@", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    worksheet.Cells[rowCounter, 9] = "Ticket";
                    FormatCells(worksheet.Columns[9], new ExcelFormatOption() { NumberFormat = "#0", HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft });

                    FormatCells(worksheet.Range[ConvertToColumn(1) + rowCounter + ":" + ConvertToColumn(9) + rowCounter], new ExcelFormatOption() { IsBold = true, IsUnderLine = true });

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
                SalespersonGroupId = salespersonGroupId
            };

        }

        private void FormatCells(Microsoft.Office.Interop.Excel.Range range, ExcelFormatOption excelFormatOption)
        {
            if (excelFormatOption.StyleName != null)
                range.Style = excelFormatOption.StyleName;
            if (excelFormatOption.NumberFormat != null)
                range.NumberFormat = excelFormatOption.NumberFormat;
  
            if (excelFormatOption.MergeCells)
                range.Merge();
            
            range.Font.Bold = excelFormatOption.IsBold;
            range.Font.Underline = excelFormatOption.IsUnderLine;
            range.HorizontalAlignment = excelFormatOption.HorizontalAlignment;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = excelFormatOption.BorderTopLineStyle;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = excelFormatOption.BorderBottomLineStyle;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = excelFormatOption.BorderLeftLineStyle;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = excelFormatOption.BorderRightLineStyle;

            range.Interior.Pattern = 1; //solid
            range.Interior.PatternColorIndex = -4105; //automatic
            switch (excelFormatOption.FillColor)
            {
                case ExcelColor.Black:
                    range.Interior.ThemeColor = 2;
                    range.Interior.TintAndShade = 0;
                    break;
                case ExcelColor.LightGray5:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.0499893185216834;
                    break;
                case ExcelColor.LightGray15:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.149998474074526;
                    break;
                case ExcelColor.LightGray25:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.249977111117893;
                    break;
                case ExcelColor.LightGray35:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.349986266670736;
                    break;
                case ExcelColor.LightOrange:
                    range.Interior.ThemeColor = 10;
                    range.Interior.TintAndShade = 0.399975585192419;
                    break;
                case ExcelColor.White:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = 0;
                    break;
            }

            switch (excelFormatOption.TextColor)
            {
                case ExcelColor.Black:
                    range.Font.ThemeColor = 2;
                    range.Font.TintAndShade = 0;
                    break;
                case ExcelColor.LightGray5:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.0499893185216834;
                    break;
                case ExcelColor.LightGray15:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.149998474074526;
                    break;
                case ExcelColor.LightGray25:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.249977111117893;
                    break;
                case ExcelColor.LightGray35:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.349986266670736;
                    break;
                case ExcelColor.LightOrange:
                    range.Font.ThemeColor = 10;
                    range.Font.TintAndShade = 0.399975585192419;
                    break;
                case ExcelColor.White:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = 0;
                    break;
            }

            range.ApplyOutlineStyles();

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

        private List<SalespersonGroup> BuildSalespersonGroup(List<Dictionary<string, object>> results)
        {
            List<SalespersonGroup> salespersonGroups = new List<SalespersonGroup>();

            foreach (Dictionary<string, object> result in results)
            {
                SalespersonGroup salespersonGroup = new SalespersonGroup();

                salespersonGroup.SalespersonGroupsId = Int32.Parse(result["salespersons_groups_id"].ToString());
                salespersonGroup.WorksheetName = result["worksheet_name"].ToString();
                salespersonGroup.SalespersonName = result["salesperson_name"].ToString();
                salespersonGroup.TerritoriesId = Int32.Parse(result["territories_id"].ToString());
                salespersonGroup.BARCForExcelStoredProcedure = result["barc_for_excel_stored_procedure"].ToString();
                salespersonGroup.SalespersonCount = Int32.Parse(result["salesperson_count"].ToString());

                salespersonGroups.Add(salespersonGroup);
            }

            return salespersonGroups;
        }

        private void RunSnapshotSprocs(CommissionRecord commissionRecord, CommissionCreateTypes createType, Int64 commissionsRecreateId, Int64 snapshotId, Dictionary<string, string> salespersons)
        {
            //only execute if we are recreating for a salesperson
            if (createType == CommissionCreateTypes.RecreateForSalesperson)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Territories",
                                        new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@pintStructuresID", commissionRecord.StructuresId));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Salespersons_Groups",
                                        new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                        new SqlParameter("@pintSnapshotsID", snapshotId));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Salespersons",
                                        new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                        new SqlParameter("@pintSnapshotsID", snapshotId));

            }

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Accounts",
                                    new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                    new SqlParameter("@pintSnapshotsID", snapshotId),
                                    new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                    new SqlParameter("@pintCommissionsMonth", commissionRecord.Month));

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Noncommissions",
                                    new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                    new SqlParameter("@pintSnapshotsID", snapshotId),
                                    new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                    new SqlParameter("@pintCommissionsMonth", commissionRecord.Month));

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Chargebacks",
                                    new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                    new SqlParameter("@pintSnapshotsID", snapshotId),
                                    new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                    new SqlParameter("@pintCommissionsMonth", commissionRecord.Month));

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Nonworking_Dates",
                                        new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                        new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate));

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Draw_Per_Days",
                                        new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                        new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate));

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Update_Snapshots_Salespersons_Performance_Goal_Percentage",
                                        new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                        new SqlParameter("@pintCommissionsMonth", commissionRecord.Month));

            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Strategies",
                                        new SqlParameter("@pintCommissionsRecreateID", commissionsRecreateId),
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                        new SqlParameter("@pintCommissionsMonth", commissionRecord.Month));

            foreach (var salesperson in salespersons)
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Playbook_Groups",
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key),
                                        new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                        new SqlParameter("@pintCommissionsMonth", commissionRecord.Month));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Playbook_Print_Division_Descriptions",
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, commissionRecord.PlaybookForBARCUpdateStoredProcedure,
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                        new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Product_Groups",
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key),
                                        new SqlParameter("@pintCommissionsYear", commissionRecord.Year),
                                        new SqlParameter("@pintCommissionsMonth", commissionRecord.Month));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Insert_Snapshots_Product_Data_Mining_Descriptions",
                                            new SqlParameter("@pintSnapshotsID", snapshotId),
                                            new SqlParameter("@pvchrSalesperson", salesperson.Key));


                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Update_Snapshots_Product_Groups_Product",
                                            new SqlParameter("@pintSnapshotsID", snapshotId),
                                            new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                            new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                            new SqlParameter("@pvchrSalesperson", salesperson.Key));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Update_Snapshots_Product_Groups_Menu_Mania",
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                        new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key));

                ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Update_Snapshots_Product_Groups_New_Business",
                                        new SqlParameter("@pintSnapshotsID", snapshotId),
                                        new SqlParameter("@psdatCommissionsMonthStartDate", commissionRecord.MonthStartDate),
                                        new SqlParameter("@psdatCommissionsEndDate", commissionRecord.EndDate),
                                        new SqlParameter("@pvchrSalesperson", salesperson.Key));
            }

        }

        private void TakeSnapshot(Int64 commissisionRecreateId, string tableName)
        {
            ExecuteNonQuery(DatabaseConnectionStringNames.Commissions, "dbo.Proc_Copy_Between_Snapshots",
                                        new SqlParameter("@pintCommissionsRecreateID", commissisionRecreateId),
                                        new SqlParameter("@pvchrTableName", tableName));
        }

        /// <summary>
        /// Validate the execute of a stored procedure that run during the recreate commmission process
        /// </summary>
        /// <param name="comm">Command to be executed</param>
        /// <param name="message">Log message prefix</param>
        /// <returns></returns>
        private bool ValidateProcedure(Dictionary<string, object> result, string message)
        {
            if (result != null)
            {
                WriteToJobLog(JobLogMessageType.WARNING, message + " by " + result["processing_by"].ToString() + " at " +
                                    String.Format("{0:MM/dd/yyyy hh:mm tt}", (DateTime)result["processing_date_time"]));
                return false;
            }

            return true;
        }

    }
}

