using BSJobBase;
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

namespace SuppliesWorkload
{
    public class Job : JobBase
    {
     //   public string GroupName { get; set; }

      //  private DatabaseConnectionStringNames VersionSpecificConnectionString { get; set; }

        public override void SetupJob()
        {
            JobName = "Supplies Workload";
            JobDescription = "Performs a bulk insert from a set of pipe delimited files into a work (staging) database";
            AppConfigSectionName = "SuppliesWorkload";

            //switch (GroupName)
            //{
            //    case "A":
            //        VersionSpecificConnectionString = DatabaseConnectionStringNames.PBSDumpAWorkLoad;
            //        break;
            //    case "B":
            //        VersionSpecificConnectionString = DatabaseConnectionStringNames.PBSDumpBWork;
            //        break;
            //    case "C":
            //        VersionSpecificConnectionString = DatabaseConnectionStringNames.PBSDumpCWork;
            //        break;
            //}

           // WriteToJobLog(JobLogMessageType.INFO, $"Group Name: {GroupName}");
        }

        public override void ExecuteJob()
        {
            try
            {
                //check for any touch files before executing
                bool touchFileFound = false;

                List<string> files = new List<string>();

                if (Directory.Exists(GetConfigurationKeyValue("TouchFileDirectory")))
                {
                    files = Directory.GetFiles(GetConfigurationKeyValue("TouchFileDirectory"), "dumpcontrol*.touch").ToList();

                    if (files != null && files.Count() > 0)
                    {
                        foreach (string file in files)
                        {
                            touchFileFound = true;


                            if (bool.Parse(GetConfigurationKeyValue("DeleteFlag")) == true)
                                File.Delete(file);
                        }
                    }
                }

                if (touchFileFound)
                {
                    //get the input files that are ready for processing
                    string inputDirectory = GetConfigurationKeyValue("InputDirectory");
                    files = Directory.GetFiles($"{inputDirectory}\\", "dumpcontrol*.timestamp").ToList();

                    if (files != null && files.Count() > 0)
                    {
                        foreach (string file in files)
                        {
                            FileInfo fileInfo = new FileInfo(file);

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.SuppliesWorkLoad, "dbo.Proc_Select_BN_Loads_DumpControl_If_Processed",
                                                                    new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                    new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                            if (previouslyLoadedFile == null)
                            {
                                //make sure the file is no longer being edited
                                if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                                {
                                    WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                    InsertLoad(fileInfo);
                                }
                            }
                            //else
                            //{
                            //    ExecuteNonQuery(VersionSpecificConnectionString, "Proc_Insert_Loads_Not_Loaded",
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
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void InsertLoad(FileInfo fileInfo)
        {

            string timeStampFileContents = File.ReadAllText(fileInfo.FullName).Replace("\n", "");

            WriteToJobLog(JobLogMessageType.INFO, $"Dump control's timestamp = " + timeStampFileContents);


            //create load record
            Int32 loadsId = 0;
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.SuppliesWorkLoad, "dbo.Proc_Insert_BN_Loads_DumpControl",
                            new SqlParameter("@pdatTimeStamp", fileInfo.LastWriteTime),
                            new SqlParameter("@pvchrOriginalDir", fileInfo.DirectoryName),
                            new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                            new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                            new SqlParameter("@pvchrUserName", Environment.UserName),
                            new SqlParameter("@pvchrComputerName", Environment.MachineName),
                            new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();

            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads Dump Control ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Update_BN_Loads_DumpControl_BNTimeStamp",
                                 new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                 new SqlParameter("@pvchrBNTimeStamp", timeStampFileContents));

            ProcessFile(fileInfo.FullName, loadsId, DateTime.Parse(timeStampFileContents));

        }

        private void ProcessFile(string fileName, Int32 loadsId, DateTime dumpControlTimeStamp)
        {
            //get a handle to the dumpcontrol*.data file. This file acts as the master list of each file to import
            FileInfo fileInfo = new FileInfo(fileName.Replace(".timestamp", ".data"));

            WriteToJobLog(JobLogMessageType.INFO, $"Reading {fileInfo.Name}");

            List<Dictionary<string, object>> tables = new List<Dictionary<string, object>>();
            //    bool pbsDumpFileVersion = false;
            //   bool suppliedDumpFileVersion = false;

            if (fileInfo.Length > 0)
            {
                List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

                foreach (string line in fileContents)
                {
                    List<string> segments = line.Split('|').ToList();

                    Dictionary<string, object> table = new Dictionary<string, object>();

                    table.Add("LoadsTableID", 0);
                    table.Add("GroupNumber", segments[1]); //always 1
                    table.Add("FromDate", segments[2]);
                    table.Add("TableName", segments[3]);
                    table.Add("FileNameWithoutExtension", segments[4]);
                    table.Add("UpdateTranNumberFileAfterSuccessfulPopulate", false);
                    table.Add("UpdateTranDateAfterSuccessfulPopulate",false);
                    table.Add("TableDumpStartDateTime", segments[5]);

                    tables.Add(table);

                }
            }

            if (tables.Count > 0)
            {

                //This sproc gets "populate immediately" flag for each group.
                Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Select_BN_Groups",
                                                                new SqlParameter("@pintGroupNumber", tables[0]["GroupNumber"])).FirstOrDefault();

                if (result == null)
                {
                    WriteToJobLog(JobLogMessageType.ERROR, $"Group number ({tables[0]["GroupNumber"]} not found)");
                    return;
                }

                DateTime fromDate;
                if (DateTime.TryParse(tables[0]["FromDate"].ToString(), out fromDate))
                    WriteToJobLog(JobLogMessageType.INFO, $"For group number {tables[0]["GroupNumber"]} , dump's from date {tables[0]["FromDate"]}");
                else
                    WriteToJobLog(JobLogMessageType.INFO, $"For group number {tables[0]["GroupNumber"]} , all records selected");

                bool populateImmediatelyAfterLoad = bool.Parse(result["populate_immediately_after_load_flag"].ToString()); 

                bool atleastOneWorkToLoad = false;

                if (populateImmediatelyAfterLoad) //always true
                {
                    atleastOneWorkToLoad = true;

                    ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Update_BN_Loads_DumpControl_Group_Number",
                                                    new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                                    new SqlParameter("@pintGroupNumber", tables[0]["GroupNumber"]));
                }
                

                //Here is where the actual data import takes place, via a bulk insert.
                List<string> filesToDelete = new List<string>();

                filesToDelete.Add(fileName); //delete dumpcontrol.timestamp file
                filesToDelete.Add(fileInfo.FullName); //delete dumpcontrol.data file
                filesToDelete.Add(fileInfo.FullName.Replace(".data", ".heading")); //delete dumpcontrol.heading file

                if (atleastOneWorkToLoad)
                {
                    string abbreviation = GetConfigurationKeyValue("Abbreviation");
                    string bulkInsertDirectory = GetConfigurationKeyValue("OutputDirectory") + abbreviation + "\\" + DateTime.Now.ToString("yyyyMMddHHmmsstt") + "\\";
                    Directory.CreateDirectory(bulkInsertDirectory);
                    Directory.CreateDirectory(bulkInsertDirectory + "Config\\");
                    Directory.CreateDirectory(bulkInsertDirectory + "Data\\");

                    WriteToJobLog(JobLogMessageType.INFO, $"Bulk insert related files will be created in {bulkInsertDirectory}Config\\");
                    WriteToJobLog(JobLogMessageType.INFO, $".data files will be copied to {bulkInsertDirectory}Data\\");

                    foreach (Dictionary<string, object> table in tables)
                    {
                        filesToDelete.AddRange(ImportTable(table, fileInfo, loadsId, bulkInsertDirectory, populateImmediatelyAfterLoad, dumpControlTimeStamp, tables));
                    }
                }

                //create workload touch file
                // CreateWorkloadTouchFile(fileInfo.Name);

                //delete files
                if (bool.Parse(GetConfigurationKeyValue("DeleteFlag")) == true)
                    DeleteFiles(filesToDelete);

                if (!populateImmediatelyAfterLoad)
                    ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, "dbo.Proc_Update_BN_Loads_DumpControl_Load_Successful_Flag", new SqlParameter("@pintLoadsDumpControlID", loadsId));

            }
        }

        private List<string> ImportTable(Dictionary<string, object> table, FileInfo fileInfo, Int32 loadsId, string bulkInsertDirectory, bool populateImmediatelyAfterLoad, DateTime dumpControlTimeStamp, List<Dictionary<string, object>> tables)
        {
            List<string> filesToDelete = new List<string>();

            string errorFile = fileInfo.DirectoryName + table["FileNameWithoutExtension"] + ".error";

            if (File.Exists(errorFile))
            {
                WriteToJobLog(JobLogMessageType.ERROR, $"Table {table["TableName"]}: .error file ({errorFile}) exists");

                //add file to list of files to delete 
                filesToDelete.Add(errorFile);

                string errorContents = File.ReadAllText(errorFile);
                throw new Exception($"Table {table["TableName"]}: Error in dump from Circulation: {errorContents}");

            }

            string timeStampFile = fileInfo.DirectoryName + "\\" + table["FileNameWithoutExtension"] + ".timestamp";

            //add file to list of files to delete
            filesToDelete.Add(timeStampFile);

            WriteToJobLog(JobLogMessageType.INFO, $"Verifying {timeStampFile} ");

            string timeStampFileContents = File.ReadAllText(timeStampFile).Replace("\n", "");

            DateTime timeStampDate;
            if (!DateTime.TryParse(timeStampFileContents, out timeStampDate))
                throw new Exception($"Unable to determine table's timestamp ({timeStampFileContents}) for table");
            else if (dumpControlTimeStamp != timeStampDate)
                throw new Exception($"Table's timestamp ({timeStampDate.ToString()}) does not match dump control's timestamp ({dumpControlTimeStamp}) for table {table["TableName"]}");

            WriteToJobLog(JobLogMessageType.INFO, $"Preparing to load {table["TableName"]}");

            Dictionary<string, object> result = null;

            if (Int32.Parse(table["LoadsTableID"].ToString()) == 0)
            {
                result = ExecuteSQL(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Insert_BN_Loads_Tables",
                                                                    new SqlParameter("@pvchrTableName", table["TableName"].ToString()),
                                                                    new SqlParameter("@pbintLoadsDumpControlID", loadsId),
                                                                    new SqlParameter("@pvchrDirectory", fileInfo.DirectoryName),
                                                                    new SqlParameter("@pvchrFile", fileInfo.Name),
                                                                    new SqlParameter("@pdatFileLastModified", fileInfo.LastWriteTime)).FirstOrDefault();
                table["LoadsTableID"] = result["loads_tables_id"];

            }
            else
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Update_BN_Loads_Tables",
                                                                new SqlParameter("@pbintLoadsTablesID", Int32.Parse(table["LoadsTableID"].ToString())),
                                                                new SqlParameter("@pvchrDirectory", fileInfo.DirectoryName),
                                                                new SqlParameter("@pvchrFile", fileInfo.Name),
                                                                new SqlParameter("@pdatFileLastModified", fileInfo.LastWriteTime));
            }



            WriteToJobLog(JobLogMessageType.INFO, $"Clearing {table["TableName"].ToString()} table for dump control's timestamp ({timeStampDate})");

            ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, CommandType.Text, $"DELETE FROM {table["TableName"].ToString()} WHERE BNTimeStamp = '{timeStampDate}'");

            string headerFile = fileInfo.DirectoryName + "\\" + table["FileNameWithoutExtension"] + ".heading";

            //add file to list of files to delete
            filesToDelete.Add(headerFile);

            WriteToJobLog(JobLogMessageType.INFO, $"Reading {headerFile}");

            List<string> fileContents = File.ReadAllLines(headerFile).ToList();
            List<Dictionary<string, object>> columnDefinitions = new List<Dictionary<string, object>>();

            foreach (string line in fileContents)
            {
                List<string> segments = line.Split('|').ToList();

                foreach (string segment in segments)
                {
                    Dictionary<string, object> dictionary = new Dictionary<string, object>();

                    dictionary.Add("ColumnIndex", 0); //this will get updated in the next loop
                    dictionary.Add("FieldLength", 0); //this will get updated in the next loop
                    dictionary.Add("ColumnName", segment);

                    columnDefinitions.Add(dictionary);
                }
            }

            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.SuppliesWorkLoad, CommandType.Text,
                                                                    $"SELECT ORDINAL_POSITION, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_CATALOG = 'Supplies_Work' AND TABLE_NAME = @TableName",
                                                                    new SqlParameter("@TableName", table["TableName"].ToString()));

            int loopCounter = 0;
            foreach (Dictionary<string, object> column in results)
            {
                loopCounter++;

                foreach (Dictionary<string, object> columnDefinition in columnDefinitions)
                {
                    if (columnDefinition["ColumnName"].ToString().ToLower() == column["COLUMN_NAME"].ToString().ToLower())
                    {
                        columnDefinition["ColumnIndex"] = loopCounter;

                        switch (column["DATA_TYPE"].ToString())
                        {
                            case "varchar":
                                //      columnDefinition["FieldLength"] = (Convert.ToInt32(column["CHARACTER_MAXIMUM_LENGTH"]) == -1 ? 8000 : column["CHARACTER_MAXIMUM_LENGTH"]);
                                break;
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

            string countFile = fileInfo.DirectoryName + "\\" + table["FileNameWithoutExtension"] + ".count";

            WriteToJobLog(JobLogMessageType.INFO, $"Reading {countFile}");

            Int64 recordCount = Int64.Parse(File.ReadAllText(countFile).ToString());

            ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Update_BN_Loads_Tables_Load_Data_Rows_Copied",
                                        new SqlParameter("@pintLoadsTablesID", table["LoadsTableID"]),
                                        new SqlParameter("@pintDataRowsCopied", recordCount));

            //add file to list of files to delete
            filesToDelete.Add(countFile);

            string bulkInsertErrorFile = bulkInsertDirectory + "Config\\" + table["TableName"].ToString() + ".error";
            string bulkInsertFormatFile = bulkInsertDirectory + "Config\\" + table["TableName"].ToString() + ".format";

            WriteToJobLog(JobLogMessageType.INFO, $"Creating {bulkInsertFormatFile}");

            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.AppendLine("9.0");
            stringBuilder.AppendLine((columnDefinitions.Count()).ToString());

            loopCounter = 1;

            foreach (Dictionary<string, object> columnDefinition in columnDefinitions)
            {
                Int32 columnIndex = Convert.ToInt32(columnDefinition["ColumnIndex"].ToString());
                stringBuilder.AppendLine($"{PadField(loopCounter.ToString(), 8)}SQLCHAR       0       {PadField(columnDefinition["FieldLength"].ToString(), 8)}\"{PadField((loopCounter == columnDefinitions.Count() ? @"\n" : "|") + "\"", 9) }{PadField(columnIndex == 0 ? "0" : columnIndex.ToString(), 6)}{PadField(columnDefinition["ColumnName"].ToString(), 39)}\"\"");

                loopCounter++;
            }

            File.WriteAllText(bulkInsertFormatFile, stringBuilder.ToString());

            string bulkInsertDataFile = table["FileNameWithoutExtension"].ToString() + ".data";

            WriteToJobLog(JobLogMessageType.INFO, $"Copying {bulkInsertDataFile} from {fileInfo.DirectoryName} to {bulkInsertDirectory} Data\\ ");

            string originalDataFile = fileInfo.DirectoryName + "\\" + bulkInsertDataFile;

            bulkInsertDataFile = bulkInsertDirectory + "Data\\" + bulkInsertDataFile;

            File.Copy(originalDataFile, bulkInsertDataFile);


            filesToDelete.Add(originalDataFile);   //add original data file to list of files to delete
            filesToDelete.Add(bulkInsertDataFile);  //add copied data file to list of files to delete

            WriteToJobLog(JobLogMessageType.INFO, $"Performing bulk insert import of {table["TableName"].ToString()} using trusted connection");
            WriteToJobLog(JobLogMessageType.INFO, $"Original data file = {originalDataFile}");
            WriteToJobLog(JobLogMessageType.INFO, $"Copied data file = {bulkInsertDataFile}");
            WriteToJobLog(JobLogMessageType.INFO, $"Error file = {bulkInsertErrorFile}");
            WriteToJobLog(JobLogMessageType.INFO, $"Format file = {bulkInsertFormatFile}");

            ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, CommandType.Text, $"BULK INSERT {table["TableName"].ToString()} FROM '{bulkInsertDataFile}' WITH (FORMATFILE='{bulkInsertFormatFile}', ERRORFILE='{bulkInsertErrorFile}')");

            WriteToJobLog(JobLogMessageType.INFO, $"Checking status of bulk insert import");

            if (File.Exists(bulkInsertErrorFile))
                throw new Exception($"Error in bulk insert. Check error file {bulkInsertErrorFile} for details");

            WriteToJobLog(JobLogMessageType.INFO, $"Deleting ignored record (last record), if read by bulk insert");

            ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, CommandType.Text, $"DELETE FROM {table["TableName"].ToString()} WHERE BNTimeStamp = '{timeStampFileContents}' AND IgnoredRecordFlag = 1");

            WriteToJobLog(JobLogMessageType.INFO, "Reading last record sequence");

            result = ExecuteSQL(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Select_RecordSequence_Maximum",
                                                             new SqlParameter("@pvchrTableName", table["TableName"].ToString()),
                                                             new SqlParameter("@pvchrBNTimeStamp", timeStampFileContents)).FirstOrDefault();

            Int64 recordSequenceMax = result["RecordSequence_maximum"].ToString() == "" ? 0 : Int64.Parse(result["RecordSequence_maximum"].ToString());

            if (recordSequenceMax == recordCount)
                WriteToJobLog(JobLogMessageType.INFO, $".count file & database both contain the same number of data records ({recordSequenceMax})");
            else
            {
                string message = $".count file ({recordCount}) differs from database count ({recordSequenceMax})";
                WriteToJobLog(JobLogMessageType.WARNING, message);
                throw new Exception(message);
            }


            if (populateImmediatelyAfterLoad) //for PBSDumpB and PBSDumpC only            
                PopulateTable(table["TableName"].ToString(), Int64.Parse(table["LoadsTableID"].ToString()), tables);


            ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, "Proc_Update_BN_Loads_Tables_Load_Successful_Flag", new SqlParameter("@pintLoadsTablesID", table["LoadsTableID"].ToString()));

            return filesToDelete;

        }

        private void DeleteFiles(List<string> files)
        {
            WriteToJobLog(JobLogMessageType.INFO, "Deleting load files");

            foreach (string file in files)
            {
                WriteToJobLog(JobLogMessageType.INFO, $"Attempting to delete {file}");
                File.Delete(file);
                WriteToJobLog(JobLogMessageType.INFO, $"{file} deleted");
            }
        }

        private void PopulateTable(string tableName, Int64 loadsTableId, List<Dictionary<string, object>> tables)
        {
            WriteToJobLog(JobLogMessageType.INFO, $"{tableName} populating");

            ExecuteNonQuery(DatabaseConnectionStringNames.SuppliesWorkLoad, $"Proc_Populate_{tableName}",
                                        new SqlParameter("@pbintLoadsTablesID", loadsTableId));

            WriteToJobLog(JobLogMessageType.INFO, $"{tableName} successful");

        }


        private string PadField(string value, int length)
        {
            if (value.Length > length)
                return value.Substring(0, length);
            else if (value.Length < length)
                return value.PadRight(length);
            else
                return value;

        }
    }
}
