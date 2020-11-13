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

namespace CircDumpWorkLoad
{
    public class Job : JobBase
    {
        //steps for job
        //1. Check for a new touch file at \\circfs\backup\circdump\touch\. This file gets extracted last by the UnzipNewscycleExportFiles job, ensuring that the rest of the batch of files are ready for processing. 
        //2. Checks \\circfs\backup\circdump\data\<groupNumber>\ for any dumpcontrol*.timestamp
        //3. For each file found, check to see if it was previously loaded 
        //4. If a file was found, create a record in BN_Loads_DumpControl table(this acts similar to the Loads table in other jobs)
        //5. Parses the dumpcontrol*.data file for a list of the files to import 
        //6. Deletes the touch file at \\Omaha\DumpTouch\CircDump\Table\<tableName>.successful(does this really ever exist?)
        //7. Create a new folder at \\Omaha\DumpTouch\CircDump\Group\<groupNumber>
        //8. Deleted the touch file at \\Omaha\DumpTouch\CircDump\Group\<groupNumber>   (does this really ever exist?)
        //9. Create a loads record for each file to be processed
        //10. Create a bulk insert directory at \\Omaha\BulkInsertFromCirc\CircDump_Work_Load_<groupNumber>\<timestamp>
        //11. Create bulk insert config directory at \\Omaha\BulkInsertFromCirc\CircDump_Work_Load_<groupNumber>\<timestamp>\Config
        //12. Create bulk insert config directory at \\Omaha\BulkInsertFromCirc\CircDump_Work_Load_<groupNumber>\<timestamp>\Data
        //13. Check for an error file if one exists \\circfs\backup\circdump\data\<groupNumber>\<tableName>.error. Exit and throw exception if one is found
        //14. Parse the file's matching timestamp file \\circfs\backup\circdump\data\<groupNumber>\<tableName>.timestamp. Make sure this matches the the timestamp in the dumpcontrol*.timestamp file
        //15. Delete any records from the destination table with a matching timestamp
        //16. Read in matching file specific header file to get a list of the column names 
        //17. Query the database for a list of column names to build the field lengths for neach column. This will then be used to build the bulk insert format file
        //18. Parse the count file at \\circfs\backup\circdump\data\<groupNumber>\<tableName>.count
        //19. Build bulk insert files, both the format and error files at \\Omaha\BulkInsertFromCirc\CircDump_Work_Load_<groupNumber>\<timestamp>\<tableName>.error and \\Omaha\BulkInsertFromCirc\CircDump_Work_Load_<groupNumber>\<timestamp>\<tableName>.format
        //20. Copy data file from \\circfs\backup\circdump\data\1\<tableName>.data to \\Omaha\BulkInsertFromCirc\CircDump_Work_Load_<groupNumber>\<timestamp>\Data\<tableName>.data
        //21. Run bulk insert
        //22. Check for new error file at \\Omaha\BulkInsertFromCirc\CircDump_Work_Load_<groupNumber>\<timestamp>\<tableName>.error, throw exception and exit if one is found
        //23. Remove last record from insert, since this is a control record
        //24. Check to make sure that all of the records were correctly inserted by comparing the count file and the last record inserted into the table
        //25. Creates a .successful in for the post tables load step to know which tables to update
        //26. Cleanups all related files
        //This is step 1 in the import process. Step 2 is CircDumpPopulate. Step 3 is CircDumpPostGroup. 

        public int GroupNumber { get; set; }

        public override void SetupJob()
        {
            JobName = "Circ Dump Workload";
            JobDescription = "Performs a bulk insert from a set of pipe delimited files into a work (staging) database";
            AppConfigSectionName = "CircDumpWorkload";
        }

        public override void ExecuteJob()
        {
            try
            {

                //check for any touch files before executing
                bool touchFileFound = false;

                List<string> files = new List<string>();

                if (Directory.Exists($"{GetConfigurationKeyValue("TouchFileDirectory")}"))
                {
                    files = Directory.GetFiles($"{GetConfigurationKeyValue("TouchFileDirectory")}", "dumpcontrol*.touch").ToList();

                    if (files != null && files.Count() > 0)
                    {
                        foreach (string file in files)
                        {
                            touchFileFound = true;
                            //only delete the file if it's the last group
                            if (GroupNumber == 6 && bool.Parse(GetConfigurationKeyValue("DeleteFlag")) == true)
                                File.Delete(file);
                        }
                    }
                }

                if (touchFileFound)
                {
                    //get the input files that are ready for processing
                    files = Directory.GetFiles($"{GetConfigurationKeyValue("InputDirectory")}{GroupNumber.ToString()}\\", "dumpcontrol*.timestamp").ToList();

                    if (files != null && files.Count() > 0)
                    {
                        WriteToJobLog(JobLogMessageType.INFO, $"Group Number: {GroupNumber}");

                        foreach (string file in files)
                        {
                            FileInfo fileInfo = new FileInfo(file);

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkLoad, "dbo.Proc_Select_BN_Loads_DumpControl_If_Processed",
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
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkLoad, "dbo.Proc_Insert_BN_Loads_DumpControl",
                            new SqlParameter("@pdatTimeStamp", fileInfo.LastWriteTime),
                            new SqlParameter("@pvchrOriginalDir", fileInfo.DirectoryName),
                            new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                            new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                            new SqlParameter("@pvchrUserName", Environment.UserName),
                            new SqlParameter("@pvchrComputerName", Environment.MachineName),
                            new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();

            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads Dump Control ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Update_BN_Loads_DumpControl_BNTimeStamp",
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

            if (fileInfo.Length > 0)
            {
                List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

                foreach (string line in fileContents)
                {
                    List<string> segments = line.Split('|').ToList();

                    Dictionary<string, object> table = new Dictionary<string, object>();

                    table.Add("LoadsTableID", 0);
                    table.Add("GroupNumber", segments[1]);
                    table.Add("FromDate", segments[2]);
                    table.Add("ArchiveEndingDate", segments[3]);
                    table.Add("TableName", segments[4]);
                    table.Add("FileNameWithoutExtension", segments[5]);
                    table.Add("UpdateTranNumberFileAfterSuccessfulPopulate", segments[6]);
                    table.Add("UpdateTranDateAfterSuccessfulPopulate", false);
                    table.Add("TableDumpStartDateTime", segments[7]);

                    tables.Add(table);

                }
            }

            if (tables.Count > 0)
            {
                //07/12/20 PEB - Added support for CircDump as part of the Newscycle Cloud migration.
                //Append the group number to this record so it's unique to the group of CircDump datasets
                ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Update_BN_Loads_DumpControl_OriginalFile",
                                                                new SqlParameter("@pintLoadsDumpControlID", loadsId),
                                                                new SqlParameter("@pvchrOriginalFile", fileName.Substring(fileName.LastIndexOf("\\") + 1) + "_" + GroupNumber.ToString()));


                //This sproc gets "populate immediately" flag for each group.  For CircDump this flag is 0 for all groups.
                Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Select_BN_Groups",
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

                bool populateImmediatelyAfterLoad = bool.Parse(result["populate_immediately_after_load_flag"].ToString()); //this is always false

                bool atleastOneWorkToLoad = false;

                ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Update_BN_Loads_DumpControl_Group_Number",
                        new SqlParameter("@pintLoadsDumpControlID", loadsId),
                        new SqlParameter("@pintGroupNumber", tables[0]["GroupNumber"].ToString()));


                //create the table records and return the loads_id for each
                foreach (Dictionary<string, object> table in tables)
                {
                    if (table["FileNameWithoutExtension"].ToString() != "")
                    {
                        atleastOneWorkToLoad = true;

                        result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Insert_BN_Loads_Tables",
                                             new SqlParameter("@pvchrTableName", table["TableName"]),
                                             new SqlParameter("@pbintLoadsDumpControlID", loadsId),
                                             new SqlParameter("@pvchrTableDumpStartDateTime", table["TableDumpStartDateTime"].ToString()),
                                             new SqlParameter("@pvchrFromDate", table["FromDate"].ToString()),
                                             new SqlParameter("@pvchrArchiveEndingDate", table["ArchiveEndingDate"].ToString()),
                                             new SqlParameter("@pflgUpdateTranNumberControlFileAfterPopulate", table["UpdateTranNumberFileAfterSuccessfulPopulate"].ToString())).FirstOrDefault();

                        table["LoadsTableID"] = result["loads_tables_id"].ToString();
                    }
                }

                //here is where the actual data import takes place, via a bulk insert.
                List<string> filesToDelete = new List<string>();

                filesToDelete.Add(fileName); //delete dumpcontrol.timestamp file
                filesToDelete.Add(fileInfo.FullName); //delete dumpcontrol.data file
                filesToDelete.Add(fileInfo.FullName.Replace(".data", ".heading")); //delete dumpcontrol.heading file

                if (atleastOneWorkToLoad)
                {
                    string bulkInsertDirectory = GetConfigurationKeyValue("OutputDirectory") + GetConfigurationKeyValue("Abbreviation") + GroupNumber.ToString() + "\\" + DateTime.Now.ToString("yyyyMMddHHmmsstt") + "\\";
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

                //delete files
                if (bool.Parse(GetConfigurationKeyValue("DeleteFlag")) == true)
                    DeleteFiles(filesToDelete);

                ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "dbo.Proc_Update_BN_Loads_DumpControl_Load_Successful_Flag",
                                       new SqlParameter("@pintLoadsDumpControlID", loadsId));
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

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Update_BN_Loads_Tables",
                                                            new SqlParameter("@pbintLoadsTablesID", Int32.Parse(table["LoadsTableID"].ToString())),
                                                            new SqlParameter("@pvchrDirectory", fileInfo.DirectoryName),
                                                            new SqlParameter("@pvchrFile", fileInfo.Name),
                                                            new SqlParameter("@pdatFileLastModified", fileInfo.LastWriteTime));


            WriteToJobLog(JobLogMessageType.INFO, $"Clearing {table["TableName"].ToString()} table for dump control's timestamp ({timeStampDate})");

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, CommandType.Text, $"DELETE FROM {table["TableName"].ToString()} WHERE BNTimeStamp = '{timeStampDate}'");

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

            List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkLoad, CommandType.Text,
                                                                    "SELECT ORDINAL_POSITION, COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_CATALOG = 'CircDump_Work' AND TABLE_NAME = @TableName",
                                                                    new SqlParameter("@TableName", table["TableName"].ToString()));

            int loopCounter = 0;
            foreach (Dictionary<string, object> column in results)
            {
                loopCounter++;

                foreach (Dictionary<string, object> columnDefinition in columnDefinitions)
                {
                    if (columnDefinition["ColumnName"].ToString() == column["COLUMN_NAME"].ToString())
                    {
                        columnDefinition["ColumnIndex"] = loopCounter;

                        switch (column["DATA_TYPE"].ToString())
                        {
                            case "varchar":
                                columnDefinition["FieldLength"] = (Convert.ToInt32(column["CHARACTER_MAXIMUM_LENGTH"]) == -1 ? 8000 : column["CHARACTER_MAXIMUM_LENGTH"]);
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

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Update_BN_Loads_Tables_Load_Data_Rows_Copied",
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

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, CommandType.Text, $"BULK INSERT {table["TableName"].ToString()} FROM '{bulkInsertDataFile}' WITH (FORMATFILE='{bulkInsertFormatFile}', ERRORFILE='{bulkInsertErrorFile}')");

            WriteToJobLog(JobLogMessageType.INFO, $"Checking status of bulk insert import");

            if (File.Exists(bulkInsertErrorFile))
                throw new Exception($"Error in bulk insert. Check error file {bulkInsertErrorFile} for details");

            WriteToJobLog(JobLogMessageType.INFO, $"Deleting ignored record (last record), if read by bulk insert");

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, CommandType.Text, $"DELETE FROM {table["TableName"].ToString()} WHERE BNTimeStamp = '{timeStampFileContents}' AND IgnoredRecordFlag = 1");

            WriteToJobLog(JobLogMessageType.INFO, "Reading last record sequence");

            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Select_RecordSequence_Maximum",
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

            ExecuteNonQuery(DatabaseConnectionStringNames.CircDumpWorkLoad, "Proc_Update_BN_Loads_Tables_Load_Successful_Flag", new SqlParameter("@pintLoadsTablesID", table["LoadsTableID"].ToString()));

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
