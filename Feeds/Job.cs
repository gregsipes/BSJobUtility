using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace Feeds
{
    public class Job : JobBase
    {
        public string Version { get; set; }

        public override void SetupJob()
        {
            JobName = "Feeds";
            JobDescription = "Create a data file that then gets sent out via sFTP.";
            AppConfigSectionName = "Feeds";

        }

        public override void ExecuteJob()
        {
            try
            {
                //todo:
                //            For intIndex = 0 To UBound(astrProgramOptionsMultiple)
                //    If astrProgramOptionsMultiple(intIndex) <> "" And Left$(astrProgramOptionsMultiple(intIndex), 1) <> "'" Then
                //        astrProgramOptionsEach = Split(astrProgramOptionsMultiple(intIndex), "=", , vbTextCompare)


                //        If UBound(astrProgramOptionsEach) = 1 Then
                //            If StrComp(astrProgramOptionsEach(0), "FeedTitle", vbTextCompare) = 0 Then
                //                strFeedTitle = astrProgramOptionsEach(1)
                //            End If


                //            If StrComp(astrProgramOptionsEach(0), "FileUpload", vbTextCompare) = 0 And StrComp(astrProgramOptionsEach(1), "True", vbTextCompare) = 0 Then
                //                flgFileUpload = True
                //            End If


                //            If StrComp(astrProgramOptionsEach(0), "KeepFTPLogForDebugging", vbTextCompare) = 0 And StrComp(astrProgramOptionsEach(1), "True", vbTextCompare) = 0 Then
                //                mflgKeepFTPLogForDebugging = True
                //            End If


                //            If StrComp(astrProgramOptionsEach(0), "PostProcessing", vbTextCompare) = 0 And StrComp(astrProgramOptionsEach(1), "True", vbTextCompare) = 0 Then
                //                flgPostProcessing = True
                //            End If
                //        End If
                //    End If
                //Next


                //todo: should the passphrase portion be brought over?


                CreateBuild();



                //test SFTP connection
                //  SFTP sFTP = new SFTP();

                // sFTP.OpenSession()

                //Int32 deletedFileCount = 0;

                //List<string> inputDirectories = GetConfigurationKeyValue("InputDirectories").Split(',').ToList();


                //foreach (string directory in inputDirectories)
                //{
                //    List<string> files = Directory.GetFiles(directory, "*.tmp").ToList();

                //    //delete each file
                //    foreach (string file in files)
                //    {
                //        FileInfo fileInfo = new FileInfo(file);

                //        if (fileInfo.Length == 0)
                //        {
                //            WriteToJobLog(JobLogMessageType.INFO, $"Deleting {file}");
                //            File.Delete(file);
                //        }
                //        deletedFileCount++;
                //    }
                //}

                //if (deletedFileCount > 0)
                //    WriteToJobLog(JobLogMessageType.INFO, $"{deletedFileCount} files deleted");

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }


        private void CreateBuild()
        {

            //retrieve the feed record from the database (this call is what replaces the differences for each feed in the INI files)
            //these fields could have been appeneded to the Feeds.dbo.Feeds table, but I didn't want to risk breaking anything
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.BSJobUtility, "Proc_Select_Feed",
                                                                new SqlParameter("@FeedName", Version)).FirstOrDefault();

            bool uploadFile = Convert.ToBoolean(result["UploadFile"].ToString());
            bool postPorcess = Convert.ToBoolean(result["PostProcess"].ToString());


            //retrieve the rest of the feed specific fields
            result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_Feeds",
                                                        new SqlParameter("@pvchrTitle", Version),
                                                        new SqlParameter("@pflgActiveOnly", 0),
                                                        new SqlParameter("@pvchrPassPhrase", ""),
                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name)).FirstOrDefault();

            DateTime? startDate = null;
            DateTime? endDate = null;

            if (Convert.ToBoolean(result["starting_date_field"].ToString()))
                startDate = DateTime.Now.AddDays(Convert.ToInt32(result["days_to_add"].ToString()));
            if (Convert.ToBoolean(result["ending_date_flag"].ToString()))
                endDate = DateTime.Now.AddDays(Convert.ToInt32(result["noninteractive_ending_date_days_after_starting_date"].ToString()));


            WriteToJobLog(JobLogMessageType.INFO, " Feeds ID: " + result["feeds_id"].ToString() +
                                                " Formats ID: " + result["formats_id"].ToString() +
                                                " Description: " + result["Description"].ToString() +
                                                " FTP Server: " + result["ftp_server"].ToString() +
                                                " Pub ID: " + result["pubid"].ToString() +
                                                " Sproc: " + result["stored_proc"].ToString() +
                                                " Username: " + result["user_name"].ToString());

            //Error checks and defaults.  Some fields might be blank (or just wrong); compute defaults and if there is no default, generate an error and exit.
            WriteToJobLog(JobLogMessageType.INFO, "Checking for errors...");

            string outputDirectory = result["output_directory"].ToString();
            //format_of_current_datetime_in_output_subdirectory
            if (result["format_of_current_datetime_in_output_subdirectory"].ToString() != "")
                outputDirectory += "\\" + DateTime.Now.ToString(result["format_of_current_datetime_in_output_subdirectory"].ToString()) + "\\";


            //todo: what should we do with the user definied log? Perhaps it can all be consolidated into 1 log stored in SQL
            //if (Convert.ToBoolean(result["log_in_output_directory_flag"].ToString()))
            //{
            //With gobjUserDefinedLog
            //    If mudfFeeds.strFileNameOfLogInOutputDirectory = "" Then
            //        .OpenUserDefinedLog strOutputDirectory &gobjIni.AppAbbrev & ".log", gobjUtilsLocal.NetAuthenticatedUserName(gobjIni.NetAuthenticatedResource), gobjUtilsLocal.ComputerName
            //    Else
            //        .OpenUserDefinedLog strOutputDirectory &mudfFeeds.strFileNameOfLogInOutputDirectory, gobjUtilsLocal.NetAuthenticatedUserName(gobjIni.NetAuthenticatedResource), _
            //            gobjUtilsLocal.ComputerName
            //    End If

            //    gflgUserDefinedLogOpened = True
            //    .WriteToUserDefinedLog "Application " & gobjIni.AppAbbrev & " (v" & App.Major & "." & App.Minor & "." & App.Revision & ")", eniInfo
            //    .WriteToUserDefinedLog "Processing " & pstrFeedsTitle, eniInfo
            //End With
            //}


            //At this point we've successfully populated all required fields, so log a message indicating that we're now building the output file.
            if (startDate.HasValue && endDate.HasValue)
                WriteToJobLog(JobLogMessageType.INFO, $"Creating build from {startDate.ToString()} thru {endDate.ToString()} ");


            //Invoke stored procedure Proc_Insert_Builds and create a record identifying (logging) this build.
            result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Insert_Builds",
                                            new SqlParameter("@pintFeedsID", result["feeds_id"].ToString()),
                                            new SqlParameter("@pvchrUserSpecifiedStartingDate", startDate.HasValue ? startDate.Value.ToString() : ""),
                                            new SqlParameter("@pvchrUserSpecifiedEndingDate", endDate.HasValue ? endDate.Value.ToString() : ""),
                                            new SqlParameter("@pvchrStandardLogFileName", ""), //todo: should this be something?
                                            new SqlParameter("@pvchrUserDefinedLogFileName", ""), //todo: should this be something?
                                            new SqlParameter("@pvchrNetworkUserName", ""),
                                            new SqlParameter("@pvchrComputerName", "")).FirstOrDefault();


        }
    }
}
