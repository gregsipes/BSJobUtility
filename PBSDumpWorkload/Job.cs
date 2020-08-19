﻿using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace PBSDumpWorkload
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "PBS Dump Workload";
            JobDescription = "";
            AppConfigSectionName = "PBSDumpWorkload";
        }

        public override void ExecuteJob()
        {
            try
            {

                //check for any touch files. If they exist, send email with contents, then delete (is this even needed?)
                if (Directory.Exists(GetConfigurationKeyValue("DumpTouchDirectory")))
                {
                    List<string> processedTouchFiles = new List<string>();
                    List<string> touchFiles = Directory.GetFiles(GetConfigurationKeyValue("DumpTouchDirectory"), GetConfigurationKeyValue("DumpTouchFile")).ToList();

                    foreach (string file in touchFiles)
                    {
                        FileInfo fileInfo = new FileInfo(file);
                        if (fileInfo.Name != "." && fileInfo.Name != ".." && processedTouchFiles.Where(p => p == file).Count() == 0)
                        {
                            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();
                            StringBuilder stringBuilder = new StringBuilder();

                            foreach (string line in fileContents)
                            {
                                if (line.Trim() != "")
                                    stringBuilder.AppendLine(line);
                            }

                            SendMail($"{GetConfigurationKeyValue("DumpTouchDescription")} Started {fileInfo.LastWriteTime.ToString()}", stringBuilder.ToString(), false);                          


                            processedTouchFiles.Add(file);
                            File.Delete(file);
                            WriteToJobLog(JobLogMessageType.INFO, $"Deleted file {file}");
                        }
                    }
                }

                //get the input files that are ready for processing
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "dumpcontrol*.timestamp").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "dbo.Proc_Select_BN_Loads_DumpControl_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                        if (previouslyLoadedFile == null)
                        {
                            //make sure we the file is no longer being edited
                            if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                            {
                                WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                InsertLoad(fileInfo);
                            }
                        }
                        //else
                        //{
                        //    ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Insert_Loads_Not_Loaded",
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

        private void InsertLoad(FileInfo fileInfo)
        {

            WriteToJobLog(JobLogMessageType.INFO, $"Dump control's timestamp = " + fileInfo.LastWriteTime.ToString());


            //create load record
            Int32 loadsId = 0;
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "dbo.Proc_Insert_BN_Loads_DumpControl",
                            new SqlParameter("@pdatTimeStamp", fileInfo.LastWriteTime),
                            new SqlParameter("@pvchrOriginalDir", fileInfo.Directory),
                            new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                            new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                            new SqlParameter("@pvchrUserName", Environment.UserName),
                            new SqlParameter("@pvchrComputerName", Environment.MachineName),
                            new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();

            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads Dump Control ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_DumpControl_BNTimeStamp",
                                 new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                 new SqlParameter("@pvchrBNTimeStamp", fileInfo.LastWriteTime));

            ProcessFile(fileInfo, loadsId);

        }

        private void ProcessFile(FileInfo fileInfo, Int32 loadsId)
        {
            WriteToJobLog(JobLogMessageType.INFO, $"Reading {fileInfo.Name}");

            List<Dictionary<string, object>> tables = new List<Dictionary<string, object>>();
            bool pbsDumpFileVersion = false;
            bool suppliedDumpFileVersion = false;

            if (fileInfo.Length > 0)
            {
                List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

                foreach (string line in fileContents)
                {
                    List<string> segments = line.Split('|').ToList();

                    Dictionary<string, object> table = new Dictionary<string, object>();

                    if (segments[0] == "PBSDump1")
                    {
                        table.Add("LoadsTableID", 0);
                        table.Add("GroupNumber", segments[1]);
                        table.Add("FromDate", segments[2]);
                        table.Add("TableName", segments[3]);
                        table.Add("FileNameWithoutExtension", segments[4]);
                        table.Add("UpdateTranNumberFileAfterSuccessfulPopulate", false);
                        table.Add("UpdateTranDateAfterSuccessfulPopulate", segments[5]);
                        table.Add("TableDumpStartDateTime", segments[6]);

                        pbsDumpFileVersion = true;
                        suppliedDumpFileVersion = false;
                    }
                    else if (segments[0] == "CircDump1")
                    {
                        table.Add("LoadsTableID", 0);
                        table.Add("GroupNumber", segments[1]);
                        table.Add("FromDate", segments[2]);
                        table.Add("ArchiveEndingDate", segments[3]);
                        table.Add("TableName", segments[4]);
                        table.Add("FileNameWithoutExtension", segments[5]);
                        table.Add("UpdateTranNumberFileAfterSuccessfulPopulate", segments[6]);
                        table.Add("UpdateTranDateAfterSuccessfulPopulate", false);
                        table.Add("TableDumpStartDateTime", segments[7]);

                        pbsDumpFileVersion = false;
                        suppliedDumpFileVersion = false;
                    }
                    else if (segments[0] == "SuppliesDump1" | segments[0] == "SuppliesDump")
                    {
                        table.Add("LoadsTableID", 0);
                        table.Add("GroupNumber", segments[1]);
                        table.Add("FromDate", segments[2]);
                        table.Add("ArchiveEndingDate", "");
                        table.Add("TableName", segments[3]);
                        table.Add("FileNameWithoutExtension", segments[4]);
                        table.Add("UpdateTranNumberFileAfterSuccessfulPopulate", false);
                        table.Add("UpdateTranDateAfterSuccessfulPopulate", false);
                        table.Add("TableDumpStartDateTime", segments[5]);

                        pbsDumpFileVersion = false;
                        suppliedDumpFileVersion = true;
                    }
                    else
                    {
                        WriteToJobLog(JobLogMessageType.WARNING, $"DumpControl File_Version {segments[0]} not defined");
                        //todo: should we send an email? Is this ever a real case?
                        return;
                    }

                   tables.Add(table);

                    //delete touch file
                    if (File.Exists(GetConfigurationKeyValue("TableTouchDirectory") + table["TableName"] + ".successful"))
                        File.Delete(GetConfigurationKeyValue("TableTouchDirectory") + table["TableName"] + ".successful");

                    //create group folder path if doesn't exist
                    if (!Directory.Exists(GetConfigurationKeyValue("GroupTouchDirectory") + table["GroupNumber"]))
                        Directory.CreateDirectory(GetConfigurationKeyValue("GroupTouchDirectory") + table["GroupNumber"]);

                    //if the file already exists, delete it
                    if (File.Exists(GetConfigurationKeyValue("GroupTouchDirectory") + table["GroupNumber"] + "\\" + table["TableName"] + ".successful"))
                        File.Delete(GetConfigurationKeyValue("GroupTouchDirectory") + table["GroupNumber"] + "\\" + table["TableName"] + ".successful");

                }
            }

            if (tables.Count > 0)
            {
                //todo: this is for the circdump specific version
                //07/12/20 PEB - Added support for CircDump as part of the Newscycle Cloud migration.
                //Append the group number to this record so it's unique to the group of CircDump datasets
                //If(mflgIsCircDump) Then
                //gcnnSQL.Proc_Update_BN_Loads_DumpControl_OriginalFile mlngLoadsDumpControlID, gobjLoad.OriginalFileName & "_" & audfTables(0).intGroupNumber
                //End If

                //This sproc gets "populate immediately" flag for each group.  For CircDump this flag is 0 for all groups.
                Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "Proc_Select_BN_Groups",
                                                                new SqlParameter("@pintGroupNumber", tables[0]["GroupNumber"])).FirstOrDefault();

                if (result == null)
                {
                    WriteToJobLog(JobLogMessageType.ERROR, $"Group number ({tables[0]["GroupNumber"]} not found)");
                    //todo: should we send an email? Is this ever a real case?
                    return;
                }

                DateTime fromDate;
                if (DateTime.TryParse(tables[0]["FromDate"].ToString(), out fromDate))
                    WriteToJobLog(JobLogMessageType.INFO, $"For group number {tables[0]["GroupNumber"]} , dump's from date {tables[0]["FromDate"]}");
                else
                    WriteToJobLog(JobLogMessageType.INFO, $"For group number {tables[0]["GroupNumber"]} , all records selected");

                bool populateImmediatelyAfterLoad = bool.Parse(result["populate_immediately_after_load_flag"].ToString());
                bool atleastOneWorkToLoad = false;

                if (populateImmediatelyAfterLoad)
                {
                    atleastOneWorkToLoad = true;

                    if (pbsDumpFileVersion)
                    {
                        //todo: this sproc doesn't exist in the current database, so hopefully this condition is never hit
                        //ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_DumpControl_Group_Number_TranDate",
                        //                                new SqlParameter("", ""),
                        //                                new SqlParameter("", ""),
                        //                                new SqlParameter("", ""),
                        //                                new SqlParameter("", ""));

                     }
                    else
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_DumpControl_Group_Number",
                                                        new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                                        new SqlParameter("@pintGroupNumber", tables[0]["GroupNumber"].ToString()));
                    }
                }
                else
                {
                    //Update the GroupNumber in the current record in BN_Loads_DumpControl
                    if (pbsDumpFileVersion)
                    {
                        //todo: this condition must not be hit, since the old version of the coode has an extra parameter and there's no matching sproc
                        //ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_DumpControl_Group_Number",
                                //new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                //new SqlParameter("@pintGroupNumber", tables[0]["GroupNumber"]));
                    } else
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_DumpControl_Group_Number",
                                new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                new SqlParameter("@pintGroupNumber", tables[0]["GroupNumber"].ToString()));
                    }

                    //create the table records and return the loads_id for each
                    foreach (Dictionary<string, object> table in tables)
                    {
                        if (table["FileNameWithoutExtension"].ToString() != "")
                        {
                            atleastOneWorkToLoad = true;

                            if (pbsDumpFileVersion)
                            {
                              result =  ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "Proc_Insert_BN_Loads_Tables",
                                                    new SqlParameter("@pvchrTableName", table["TableName"]),
                                                    new SqlParameter("@pbintLoadsDumpControlID", loadsId),
                                                    new SqlParameter("@pvchrTableDumpStartDateTime", table["TableDumpStartDateTime"].ToString()),
                                                    new SqlParameter("@pvchrFromDate", table["FromDate"].ToString())).FirstOrDefault();
                            } 
                            else
                            {
                               result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "Proc_Insert_BN_Loads_Tables",
                                                    new SqlParameter("@pvchrTableName", table["TableName"]),
                                                    new SqlParameter("@pbintLoadsDumpControlID", loadsId),
                                                    new SqlParameter("@pvchrTableDumpStartDateTime", table["TableDumpStartDateTime"].ToString()),
                                                    new SqlParameter("@pvchrFromDate", table["FromDate"].ToString()),
                                                    new SqlParameter("@pvchrArchiveEndingDate", table["ArchiveEndingDate"].ToString()),
                                                    new SqlParameter("@pflgUpdateTranNumberControlFileAfterPopulate", table["UpdateTranNumberFileAfterSuccessfulPopulate"].ToString())).FirstOrDefault();
                            }

                            table["LoadsTableID"] = result["loads_table_id"].ToString();
                        }
                    }


                    //todo: this must not be hit, there's no matching sproc with this name
                    //if (pbsDumpFileVersion)
                    //    ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Insert_BN_DumpControl_Post_Load", new SqlParameter("", loadsId));

                    //Here is where the actual data import takes place, via a bulk insert.
                    if (atleastOneWorkToLoad)
                    {
                        string bulkInsertDirectory = GetConfigurationKeyValue("OutputDirectory") + GetConfigurationKeyValue("Abbreviation") + "\\" + DateTime.Now.ToString("yyyymmddhhnnssms") + "\\";
                        Directory.CreateDirectory(bulkInsertDirectory);
                        Directory.CreateDirectory(bulkInsertDirectory + "Config\\");
                        Directory.CreateDirectory(bulkInsertDirectory + "Data\\");

                        WriteToJobLog(JobLogMessageType.INFO, $"Bulk insert related files will be created in {bulkInsertDirectory}Config\\");
                        WriteToJobLog(JobLogMessageType.INFO, $".data files will be copied to {bulkInsertDirectory}Data\\");


                        foreach (Dictionary<string, object> table in tables)
                        {
                            ImportTable(table, fileInfo, loadsId);
                        }
                    }


                }
            }

        }

        private void ImportTable(Dictionary<string, object> table, FileInfo fileInfo, Int32 loadsId)
        {
            string errorFile = fileInfo.DirectoryName + table["FileNameWithoutExtension"] + ".error";

            if (File.Exists(errorFile))
            {
                WriteToJobLog(JobLogMessageType.ERROR, $"Table {table["TableName"]}: .error file ({errorFile}) exists");

                //todo: add file to list of files to delete

                string errorContents = File.ReadAllText(errorFile);
                throw new Exception($"Table {table["TableName"]}: Error in dump from Circulation: {errorContents}");

            }

            string timeStampFile = fileInfo.DirectoryName + table["FileNameWithoutExtension"] + ".timestamp";
            
            //todo: add file to list of files to delete

            WriteToJobLog(JobLogMessageType.INFO, $"Verifying {timeStampFile} ");

            string timeStampFileContents = File.ReadAllText(timeStampFile);

            DateTime timeStampDate;
            if (!DateTime.TryParse(timeStampFile, out timeStampDate))
                throw new Exception($"Unable to determine table's timestamp ({timeStampFileContents}) for table");
            else if (fileInfo.LastAccessTime != timeStampDate)
                throw new Exception($"Table's timestamp ({timeStampDate.ToString()}) does not match dump control's timestamp ({fileInfo.LastAccessTime}) for table {table["TableName"]}");

            WriteToJobLog(JobLogMessageType.INFO, $"Preparing to load {table["TableName"]}");

            if (Int32.Parse(table["LoadsTableID"].ToString()) == 0)
            {
                //todo: this must not get hit; the parameters in the code don't match the sproc
               Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, "Proc_Insert_BN_Loads_Tables",
                                                                new SqlParameter("@pvchrTableName", table["TableName"].ToString()),
                                                                new SqlParameter("@pbintLoadsDumpControlID", loadsId),
                                                                new SqlParameter("@pvchrTableDumpStartDateTime", table["TableDumpStartDateTime"].ToString()),
                                                                new SqlParameter("@pvchrFromDate", table["FromDate"].ToString()),
                                                                new SqlParameter("@pvchrArchiveEndingDate", table["ArchiveEndingDate"].ToString()),
                                                                new SqlParameter("@pflgUpdateTranNumberControlFileAfterPopulate", table["UpdateTranNumberFileAfterSuccessfulPopulate"].ToString())).FirstOrDefault();
                table["LoadsTableID"] = result["loads_table_id"];
            }
            else
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, "Proc_Update_BN_Loads_Tables",
                                                             new SqlParameter("@pbintLoadsTablesID", Int32.Parse(table["LoadsTableID"].ToString())),
                                                             new SqlParameter("@pvchrDirectory", fileInfo.DirectoryName),
                                                             new SqlParameter("@pvchrFile", fileInfo.Name),
                                                             new SqlParameter("@pdatFileLastModified", fileInfo.LastWriteTime));
            }

            WriteToJobLog(JobLogMessageType.INFO, $"Clearing {table["TableName"].ToString()} table for dump control's timestamp ({fileInfo.LastAccessTime})");

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWork, CommandType.Text, $"DELETE FROM {table["TableName"].ToString()} WHERE BNTimeStamp = '{fileInfo.LastAccessTime}'");

            string headerFile = fileInfo.DirectoryName + table["FileNameWithoutExtension"] + ".heading";

            //todo: add file to list of files to delete

            WriteToJobLog(JobLogMessageType.INFO, $"Reading {headerFile}");

            List<string> fileContents = File.ReadAllLines(headerFile).ToList();
            List<Dictionary<string, object>> columnDefinitions = new List<Dictionary<string, object>>();

            foreach (string line in fileContents)
            {
                List<string> segments = line.Split('|').ToList();
                Dictionary<string, object> dictionary = new Dictionary<string, object>();

                dictionary.Add("FieldLength", 0); //this will get updated in the next loop
                dictionary.Add("ColumnName", segments[0].ToString());

                columnDefinitions.Add(dictionary);
            }

            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWork, CommandType.Text,
                                                                    "SELECT ORDINAL_POSITION, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_CATALOG = 'CircDump_Work' AND TABLE_NAME = '@TableName'",
                                                                    new SqlParameter("@TableName", table["TableName"].ToString());

            foreach (Dictionary<string, object> result in results)
            {
                foreach (Dictionary<string, object> columnDefinition in columnDefinitions)
                {
                    if (columnDefinition["ColumnName"].ToString() == result["COLUMN_NAME"].ToString())
                    {
                        switch (result["DATA_TYPE"].ToString())
                        {
                            case "int":
                                columnDefinition["FieldLength"] = 12;
                                break;
                            case "bigint":
                                columnDefinition["FieldLength"] = 19;
                                break;
                            case "datetime":
                            case "smalldatetime":
                                columnDefinition["FieldLength"] = 24;
                                break;
                            case "bit":
                                columnDefinition["FieldLength"] = 1;
                                break;
                            case "money":
                                columnDefinition["FieldLength"] = 30;
                                break;
                            case "decimal":
                                columnDefinition["FieldLength"] = 41;
                                break;
                            case "tinyint":
                                columnDefinition["FieldLength"] = 5;
                                break;
                        }
                    }
                }
            }

        }
    }
}
