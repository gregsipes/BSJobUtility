﻿using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data;
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
                                                        new SqlParameter("@pflgActiveOnly", false),
                                                        new SqlParameter("@pvchrPassPhrase", ""),
                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name)).FirstOrDefault();

            DateTime? startDate = null;
            DateTime? endDate = null;


            if (Convert.ToBoolean(feed["starting_date_flag"].ToString()))
                startDate = DateTime.Now.AddDays(Convert.ToInt32(feed["days_to_add"].ToString()));
            if (Convert.ToBoolean(feed["ending_date_flag"].ToString()))
                endDate = startDate.Value.AddDays(Convert.ToInt32(feed["noninteractive_ending_date_days_after_starting_date"].ToString()));

            //todo: remove, test code only
            startDate = new DateTime(2020, 11, 27);


            WriteToJobLog(JobLogMessageType.INFO, " Feeds ID: " + feed["feeds_id"].ToString() +
                                                " Formats ID: " + feed["formats_id"].ToString() +
                                                " Description: " + feed["description"].ToString() +
                                                " FTP Server: " + feed["ftp_server"].ToString() +
                                                " Pub ID: " + feed["pubid"].ToString() +
                                                " Sproc: " + feed["stored_proc"].ToString() +
                                                " Username: " + feed["user_name"].ToString());

            //Error checks and defaults.  Some fields might be blank (or just wrong); compute defaults and if there is no default, generate an error and exit.
            WriteToJobLog(JobLogMessageType.INFO, "Checking for errors...");

            string outputDirectory = feed["output_directory"].ToString();
            //format_of_current_datetime_in_output_subdirectory
            if (feed["format_of_current_datetime_in_output_subdirectory"].ToString() != "")
                outputDirectory += "\\" + DateTime.Now.ToString(feed["format_of_current_datetime_in_output_subdirectory"].ToString()) + "\\";


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
            // if (startDate.HasValue && endDate.HasValue)
            WriteToJobLog(JobLogMessageType.INFO, $"Creating build from {startDate.ToString()} thru {endDate.ToString() ?? ""} ");


            //Invoke stored procedure Proc_Insert_Builds and create a record identifying (logging) this build.
            result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Insert_Builds",
                                            new SqlParameter("@pintFeedsID", feed["feeds_id"].ToString()),
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

            Int64 userSerialNumber = 0;
            //Aging summary (special case, only when this build's aging_summary_flag is set to 1)
            if (Convert.ToBoolean(feed["aging_summary_flag"].ToString()))
            {
                WriteToJobLog(JobLogMessageType.INFO, "Assigning userserialno for Aging Summary.");

                result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_UserSerialNos").FirstOrDefault();

                userSerialNumber = Convert.ToInt64(result["userserialno"].ToString());

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
            }

            //Invoke the appropriate stored procedure (from the build record field "stored_proc" in table Feeds)
            string parameterString = DetermineParameters(Convert.ToInt64(feed["feeds_id"].ToString()), buildId, feed["pubid"].ToString(), userSerialNumber, startDate.HasValue ? startDate.Value.ToShortDateString() : "", endDate.HasValue ? endDate.Value.ToShortDateString() : "", feed["user_name"].ToString());


            WriteToJobLog(JobLogMessageType.INFO, "Selecting data with parameters");

            //(The "mudfFeels.strStoredProc" value can be found in table Feeds, field stored_proc - IF you know the feeds_id value.
            //For Tearsheets, this would be a feeds_id = 7,
            //which translates to "Proc_Select_Tearsheets"

            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Feeds, CommandType.Text,  "EXEC " + feed["stored_proc"].ToString() + " " + parameterString).ToList();

            if (results.Count() <= 0)
            {
                WriteToJobLog(JobLogMessageType.INFO, "No data selected for this feed.");

                //throw an exception if the feed's flag is set to true
                if (Convert.ToBoolean(feed["error_if_no_data_selected_flag"].ToString()))
                    throw new Exception($"No data selected for feed");

                ExecuteNonQuery(DatabaseConnectionStringNames.Feeds, "Proc_Update_Builds_End",
                                    new SqlParameter("@pintBuildsID", buildId));

                return;
            }


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
                    //Execute this logic for any field-delimited outputs (delimited with a comma, pipe, etc.)
                    if (Convert.ToBoolean(format["headings_flag"].ToString()))
                    {
                        foreach (Dictionary<string, object> field in fields)
                        {
                            stringBuilder.Append(format["quote_character"].ToString() + field["output_field"].ToString() + format["quote_character"].ToString());
                        }
                    }

                    //For EVERY record in the dataset,
                    //Convert each value to a string, contatenate it with the appropriate delimiter, and output it
                    //   to the output file.
                    // Exit this loop on any conversion error.
                    // Int64 dataRowCounter = 0;
                    foreach (Dictionary<string, object> dataRow in results)
                    {
                        //  dataRowCounter++;

                        //With each pass in the loop below, populate this string with a delimiter and a formatted field value
                        foreach (Dictionary<string, object> field in fields)
                        {
                            //stringBuilder.Append(FormatField(dataRow[field["source_field"].ToString()].ToString(), field["format_string"].ToString()) + format["delimiter_character"].ToString());
                            stringBuilder.Append(dataRow[field["source_field"].ToString()].ToString() + format["delimiter_character"].ToString());

                        }

                        //todo: is this needed?
                        // if (Convert.ToBoolean(feed["source_of_additional_information_in_log_files"].ToString()))


                        //Create a master list of PDF files to be FTP'd.These come from the selected build sproc(mrstSQL!file_name)
                        // and will be FTP'd during post - processing
                        if (postProcess && Convert.ToInt32(feed["post_processing_group"].ToString()) == 2)
                            filesToPostProcess.Add(dataRow["file_name"].ToString());
                    }
                }
                else if (Convert.ToBoolean(format["fixed_width_flag"].ToString()))
                {
                    //Execute this logic for fixed-width field outputs.
                    if (Convert.ToBoolean(format["headings_flag"].ToString()))
                    {
                        foreach (Dictionary<string, object> field in fields)
                        {
                            if (Convert.ToBoolean(field["left_justified_flag"].ToString()))   //this is always true
                                stringBuilder.Append(field["output_field"].ToString().PadRight(Convert.ToInt32(field["field_length"].ToString())));
                            else
                                stringBuilder.Append(field["output_field"].ToString().PadLeft(Convert.ToInt32(field["field_length"].ToString())));
                        }
                    }


                    foreach (Dictionary<string, object> dataRow in results)
                    {
                        //  dataRowCounter++;

                        //With each pass in the loop below, populate this string with a delimiter and a formatted field value
                        foreach (Dictionary<string, object> field in fields)
                        {
                            stringBuilder.Append(dataRow[field["source_field"].ToString()].ToString().PadRight(Convert.ToInt32(field["field_length"].ToString())));

                        }

                        //todo: is this needed?
                        // if (Convert.ToBoolean(feed["source_of_additional_information_in_log_files"].ToString()))


                        //Create a master list of PDF files to be FTP'd.These come from the selected build sproc(mrstSQL!file_name)
                        // and will be FTP'd during post - processing
                        if (postProcess && Convert.ToInt32(feed["post_processing_group"].ToString()) == 2)
                            filesToPostProcess.Add(dataRow["file_name"].ToString());
                    }

                }

            }

            //In table Builds, set the file_creation_end_date_time field to current date/time
            ExecuteNonQuery(DatabaseConnectionStringNames.Feeds, "Proc_Update_Builds_File_Creation_End",
                    new SqlParameter("@pintBuildsID", buildId),
                    new SqlParameter("@pintDataRecordCount", results.Count()));

            //"POST PROCESSING" is where files are transferred from the local source to the remote (FTP or SFTP) destination
            if (postProcess)
            {
                result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Update_Builds_Post_Processing_Start",
                                        new SqlParameter("@pintBuildsID", buildId)).FirstOrDefault();

                bool continueProcessingOnError = Convert.ToBoolean(result["continue_processing_if_fails_flag"]);

                // if (Convert.ToInt32(result["post_processing_group"].ToString()) == 0)

                bool successful = PostProcess(Convert.ToInt32(result["post_processing_group"].ToString()), feed, filesToPostProcess, outputFileName);

                if (!successful && !continueProcessingOnError)
                    throw new Exception("Post Process unsuccessful and continuing processing set to false");

                result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Update_Builds_Post_Processing_End",
                new SqlParameter("@pintBuildsID", buildId)).FirstOrDefault();
            }

        }

        //private string FormatField(string value, string format)
        //{
        //    //Specialized formatting routine to convert bit (or string) 1/0 into string containing
        //    //Yes/No
        //    //True/False
        //    //On/Off
        //    if (format != null)
        //    {
        //        switch (format)
        //        {
        //            case "0/1":
        //                if (value == "0")
        //                    return "0";
        //                else
        //                    return "1";
        //            case ""
        //        }
        //    }


        //    return field;

        //}

        private bool PostProcess(Int32 groupNumber, Dictionary<string, object> feed, List<string> filesToPostProcess, string outputFileName)
        {
            //there are only 2 groups, 1 and 2
            //group 1 is only for video employment ads. These ads were stopped in early 2020 but the code was carried over just in case it was needed again

            switch (groupNumber)
            {
                case 1:
                    //todo:
                    break;
                case 2:

                    if (filesToPostProcess.Count() == 0)
                    {
                        WriteToJobLog(JobLogMessageType.INFO, "No files to be ftp'd in post processing");
                        return false;
                    }

                    Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_Post_Processing_Groups",
                                                                            new SqlParameter("@pintPostPorcessingGroup", groupNumber)).FirstOrDefault();

                    //Open an FTP channel
                    WriteToJobLog(JobLogMessageType.INFO, $"FTP Server = {feed["isSFTP"].ToString()}");
                    WriteToJobLog(JobLogMessageType.INFO, $"FTP User = {feed["user_name"].ToString()}");
                    WriteToJobLog(JobLogMessageType.INFO, $"Binary FTP? = {feed["binary_flag"].ToString()}");
                    WriteToJobLog(JobLogMessageType.INFO, $"SFTP? = {feed["isSFTP"].ToString()}");
                    WriteToJobLog(JobLogMessageType.INFO, $"Remote destination directory? = {feed["put_subdirectory"].ToString()}");


                    //either ftp or stp the files
                    if (Convert.ToBoolean(feed["isSFTP"].ToString()))
                    {
                        SFTP sFTP = new SFTP(feed["ftp_server"].ToString(), feed["user_name"].ToString(), feed["Password"].ToString());

                        sFTP.OpenSession();

                        //create the destination directory if one doesn't already exist
                        if (!sFTP.CheckIfDirectoryExists(feed["put_subdirectory"].ToString()))
                        {
                            WriteToJobLog(JobLogMessageType.INFO, "Remote directory does not exist");

                            sFTP.CreateDirectory(feed["put_subdirectory"].ToString());   //todo: do we want to add looping here?
                        }

                        //Output every name on the FTP file list (that came from the list built in CreateBuild())
                        foreach (string file in filesToPostProcess)
                        {
                            sFTP.UploadFile(file, feed["put_subdirectory"].ToString(), true, true);

                            WriteToJobLog(JobLogMessageType.INFO, $"Successfully uploaded {file}");

                            //todo: should we add a retry counter?
                        }

                        WriteToJobLog(JobLogMessageType.INFO, $"Successfully uploaded {filesToPostProcess.Count()} files");



                        sFTP.CloseSession();

                    } else
                    {
                        FTP ftp = new FTP(feed["ftp_server"].ToString(), feed["user_name"].ToString(), feed["Password"].ToString());

                        //create the destination directory if one doesn't already exist
                        if (!ftp.CheckIfDirectoryExists(feed["put_subdirectory"].ToString()))
                        {
                            WriteToJobLog(JobLogMessageType.INFO, "Remote directory does not exist");

                            ftp.CreateDirectory(feed["put_subdirectory"].ToString());   //todo: do we want to add looping here?
                        }

                        //Output every name on the FTP file list (that came from the list built in CreateBuild())
                        foreach (string file in filesToPostProcess)
                        {
                            ftp.UploadFile(new System.IO.FileInfo(file), feed["put_subdirectory"].ToString());

                            WriteToJobLog(JobLogMessageType.INFO, $"Successfully uploaded {file}");

                            //todo: should we add a retry counter?
                        }

                        WriteToJobLog(JobLogMessageType.INFO, $"Successfully uploaded {filesToPostProcess.Count()} files");

                    }






                    break;
            }


            return true;



        }

        private string DetermineParameters(Int64 feedId, Int64 buildId, string pubId, Int64 userSerialNumber, string startDate, string endDate, string userName)
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
                        parameterString += "'" + GetConfigurationKeyValue("RemoteDatabaseName") + "',";
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
                        parameterString += "'" + GetConfigurationKeyValue("RemoteUserName") + "',";
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

            if (!directory.EndsWith("\\") && directory != "")
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
