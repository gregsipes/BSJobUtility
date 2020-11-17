﻿using BSJobBase;
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
            bool postProcess = Convert.ToBoolean(result["PostProcess"].ToString());


            //retrieve the rest of the feed specific fields
            Dictionary<string, object> feed = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_Feeds",
                                                        new SqlParameter("@pvchrTitle", Version),
                                                        new SqlParameter("@pflgActiveOnly", 0),
                                                        new SqlParameter("@pvchrPassPhrase", ""),
                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name)).FirstOrDefault();

            DateTime? startDate = null;
            DateTime? endDate = null;

            if (Convert.ToBoolean(result["starting_date_flag"].ToString()))
                startDate = DateTime.Now.AddDays(Convert.ToInt32(result["days_to_add"].ToString()));
            if (Convert.ToBoolean(result["ending_date_flag"].ToString()))
                endDate = startDate.Value.AddDays(Convert.ToInt32(result["noninteractive_ending_date_days_after_starting_date"].ToString()));


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
                                            new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                            new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower())).FirstOrDefault();

            Int64 buildId = Convert.ToInt64(result["builds_id"].ToString());

            //Invoke 'Proc_Select_Formats to get the correct format for this particular build.
            //formats_id is the parameter that selects for the specfic build.
            //(Example:  In the case of TSExport, formats_id = 4)

            Dictionary<string, object> format = null;
            List<Dictionary<string, object>> fields = null;

            if (Convert.ToInt32(feed["formats_id"].ToString()) != 0)
            {
                //get format
                WriteToJobLog(JobLogMessageType.INFO, "Determining feed file format.");
                format = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_Formats", new SqlParameter("@pintFormatsID", Convert.ToInt32(feed["formats_id"].ToString()))).FirstOrDefault();

                //get fields
                WriteToJobLog(JobLogMessageType.INFO, "Determining fields associated with feed file format.");
                fields = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_Fields", new SqlParameter("@pintFormatsID", Convert.ToInt32(feed["formats_id"].ToString()))).ToList();

            }


            //Aging summary (special case, only when this build's aging_summary_flag is set to 1)
            if (Convert.ToBoolean(feed["aging_summary_flag"].ToString()))
            {
                WriteToJobLog(JobLogMessageType.INFO, "Assigning userserialno for Aging Summary.");

                result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_UserSerialNos").FirstOrDefault();

                Int64 userSerialNumber = Convert.ToInt64(result["userserialno"].ToString());

                WriteToJobLog(JobLogMessageType.INFO, "Retrieving dates for Aging Summary.");

                result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_tblSites",
                                    new SqlParameter("@pvchrBWServerInstance", GetConfigurationKeyValue("RemoteServerName")),
                                    new SqlParameter("@pvchrBWDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                    new SqlParameter("@pvchrBWUserName", GetConfigurationKeyValue("RemoteUserName")),
                                    new SqlParameter("@pvchrBWPassword", GetConfigurationKeyValue("RemotePassword"))).FirstOrDefault();

                ExecuteNonQuery(DatabaseConnectionStringNames.Brainworks, "PrepareAsOfAgingSummarynew",
                                    new SqlParameter("@asofagingdate", DateTime.Now.ToShortDateString()),
                                    new SqlParameter("@current", Convert.ToDateTime(result["periodstartdate"].ToString()).ToShortDateString()),
                                    new SqlParameter("@days30", Convert.ToDateTime(result["days30"].ToString()).ToShortDateString()),
                                    new SqlParameter("@days60", Convert.ToDateTime(result["days60"].ToString()).ToShortDateString()),
                                    new SqlParameter("@days90", Convert.ToDateTime(result["days90"].ToString()).ToShortDateString()),
                                    new SqlParameter("@UserSerialno", userSerialNumber));

                //Invoke the appropriate stored procedure (from the build record field "stored_proc" in table Feeds)
                string parameterString = DetermineParameters(Convert.ToInt64(result["feeds_id"].ToString()), buildId, feed["pubid"].ToString(), userSerialNumber, startDate.Value, endDate.Value, feed["user_name"].ToString());

                WriteToJobLog(JobLogMessageType.INFO, "Selecting data with parameters");

                //(The "mudfFeels.strStoredProc" value can be found in table Feeds, field stored_proc - IF you know the feeds_id value.
                //For Tearsheets, this would be a feeds_id = 7,
                //which translates to "Proc_Select_Tearsheets"

                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "EXEC " + feed["stored_proc"].ToString() + " " + parameterString).ToList();

                //todo: are we supposed to be reusing this variable?
                buildId = Convert.ToInt64(result["builds_id"].ToString());

                if (feed["date_column_for_put_subdirectory_replacement"].ToString() != "")
                {
                    string replacementDateString = results[0][feed["date_column_for_put_subdirectory_replacement"].ToString()].ToString();
                    DateTime replacementDate;
                    if (DateTime.TryParse(replacementDateString, out replacementDate))
                    {
                        string subDirectory = feed["put_subdirectory"].ToString();
                        subDirectory = subDirectory.Replace("{dd}", replacementDate.ToString("dd"));
                        subDirectory = subDirectory.Replace("{mm}", replacementDate.ToString("MM"));
                        subDirectory = subDirectory.Replace("{yy}", replacementDate.ToString("yy"));
                        subDirectory = subDirectory.Replace("{yyyy}", replacementDate.ToString("yyyy"));

                        feed["put_subdirectory"] = subDirectory;
                    }
                }


                //Create output filename:  For Tearsheets it's in the form TSExport_YYMMDD_YYMMddhhmmss.txt
                string outputFileName = DetermineOutputFileName(outputDirectory, feed["output_file_name_prefix"].ToString(), feed["format_of_user_specified_date_in_output_file_name"].ToString(), 
                                                               feed["format_of_current_datetime_in_output_file_name"].ToString(), endDate, feed["output_file_name_extension"].ToString());

                //In table Builds, set the file_creation_start_date_time field to current date/time
                ExecuteNonQuery(DatabaseConnectionStringNames.Feeds, "Proc_Update_Builds_File_Creation_Start",
                                            new SqlParameter("@pintBuildsID", buildId),
                                            new SqlParameter("@pdatCurrent", DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt")),
                                            new SqlParameter("@pvchrDataFileName", outputFileName));

                if (outputFileName != "")
                    WriteToJobLog(JobLogMessageType.INFO, $"Creating feed file {outputFileName}");


                //When no target output filename has been specified, ONLY
                //create a master list of PDF files to be FTP'd.These come from the selected build sproc(mrstSQL!file_name)
                //and will be FTP'd during post - processing
                List<string> filesToPostProcess = new List<string>();
                if (outputFileName == "")
                {
                    foreach (Dictionary<string, object> filesToCreate in results)
                    {
                        if (postProcess && Convert.ToInt32(feed["post_processing_group"].ToString()) == 2)
                            filesToPostProcess.Add(filesToCreate["file_name"].ToString());
                    }
                }
                else
                {
                    //When a target output filename has been specified, write specific build fields from each build record to this file
                    //and then FTP the corresponding PDF file.
                    StringBuilder stringBuilder = new StringBuilder();
                   if (Convert.ToBoolean(format["delimited_flag"].ToString()))
                    {
                        if (Convert.ToBoolean(format["headings_flag"].ToString()))
                        {
                            foreach(Dictionary<string, object> field in fields)
                            {

                            }
                        }
                    }
                }

            }
        }

        private string DetermineParameters(Int64 feedId, Int64 buildId, string pubId, Int64 userSerialNumber, DateTime startDate, DateTime endDate, string userName)
        {
            //Different ad types have different fields/parameters associated with them.
            //Here is where we select the appropriate parameters based on the FeedsID
            //(which, of course, is passed in as a global parameter and not a calling parameter,
            string parameterString = "";

            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_Parameters", new SqlParameter("@pintFeedsID", feedId)).ToList();

            foreach (Dictionary<string, object> result in results)
            {
                switch (result["parameter_name"].ToString())
                {
                    case "builds_id":
                        parameterString += buildId + ",";
                        break;
                    case "bw_database":
                        parameterString += "'" + GetConfigurationKeyValue("RemoteDatabaseName") + "'," ;
                        break;
                    case "bw_server_instance":
                        parameterString += "'" + GetConfigurationKeyValue("RemoteServerName") + "',";
                        break;
                    case "ending_date":
                        parameterString += "'" + endDate + "',";
                        break;
                    case "false":
                        parameterString += "0" + ",";
                        break;
                    case "password":
                        parameterString += "'" + GetConfigurationKeyValue("RemotePassword") + "',";
                        break;
                    case "pbsdumpb_database":
                        parameterString += "'" + GetConfigurationKeyValue("PBSDumpBDatabaseName") + "',";
                        break;
                    case "pbsdumpb_server_instance":
                        parameterString += "'" + GetConfigurationKeyValue("PBSDumpBServerName") + "',";
                        break;
                    case "pubid":
                        parameterString += pubId + ",";
                        break;
                    case "starting_date":
                        parameterString += "'" + startDate + "',";
                        break;
                    case "true":
                        parameterString += "1" + ",";
                        break;
                    case "user_name":
                        parameterString += "'" + userName + "',";
                        break;
                    case "userserialno":
                        parameterString += "'" + userSerialNumber + "',";
                        break;
                }
            }

            return parameterString.TrimEnd(',');

        }

        private string DetermineOutputFileName(string directory, string prefix, string dateFormat, string outputFileDateFormat, DateTime? endDate, string extension)
        {

            if (extension == "")
                extension = ""; //does this ever get hit?
            else
                extension = "." + extension;

            if (!directory.EndsWith("\\"))
                directory += "\\";

            string outputFileName = directory + prefix;

            string dateFormatString = dateFormat.Replace("m", "M");
            string timeFormatString = outputFileDateFormat.Replace("m", "M").Replace("n", "m");

            if (dateFormatString != "" & endDate.HasValue)
                outputFileName += endDate.Value.ToString(dateFormatString);

            if (timeFormatString != "")
                outputFileName += DateTime.Now.ToString(timeFormatString);

            outputFileName += extension;


            return outputFileName;
        }
    }
}
