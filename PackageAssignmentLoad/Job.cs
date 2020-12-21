using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BSJobBase;
using System.Data.SqlClient;
using static BSGlobals.Enums;
using System.IO;
using System.Reflection;

namespace PackageAssignmentLoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "PackageAssignmentLoad";
            JobDescription = "TODO";
            AppConfigSectionName = "PackageAssignmentLoad";

        }

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "package.*").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Length > 0) //ignore empty files
                        {

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.Manifests, "dbo.Proc_Select_Packages_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                            if (previouslyLoadedFile == null)
                            {
                                //make sure we the file is no longer being edited
                                if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                                {
                                    WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                    CopyAndProcessFile(fileInfo);
                                }
                            }
                            //else
                            //{
                            //    ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Loads_Not_Loaded",
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

        private void CopyAndProcessFile(FileInfo fileInfo)
        {
            string backupFileName = GetConfigurationKeyValue("BackupDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".txt";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName, true);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Packages_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Update_Packages_Loads_Backup",
                                        new SqlParameter("@pintPackagesLoadsID", loadsId),
                                        new SqlParameter("@pstrBackupFile", backupFileName));

            //parse file and store contents
            List<string> fileContents = File.ReadAllLines(fileInfo.FullName).ToList();

            string reportDate = "";
            string mixName = "";
            string currentPreprintName = "";
            bool inFooterSection = false;
            List<string> distinctPackageNumbers = new List<string>();
            List<string> currentPagePackageNumbers = new List<string>();

            foreach (string line in fileContents)
            {
                if (line.Trim() != "")
                {
                    if (line.StartsWith("PACKAGE ASSIGNMENT REPORT -"))   //this is the header line of the file
                    {
                        List<string> segments = line.Split(' ').ToList();
                        reportDate = segments[4].Replace("PACKAGE ASSIGNMENT REPORT -", "").Trim();
                        mixName = segments[13].Replace("MIX -", "").Trim();
                    }
                    else if (line.StartsWith("PREPRINTS     |Hop"))  //this is column header line for the body of the report
                    {
                        currentPagePackageNumbers  = line.Replace("PREPRINTS     |Hop|", "").Split('|').ToList();

                        foreach (string packageNumber in currentPagePackageNumbers.Where(p => p.Trim() != ""))
                        {
                            if (!distinctPackageNumbers.Contains(packageNumber))
                            {
                                ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Packages",
                                                                new SqlParameter("@pintPackagesLoadsID", loadsId),
                                                                new SqlParameter("@pintPackageNumber", packageNumber));
                                distinctPackageNumbers.Add(packageNumber);
                            }
                        }
                    }
                    else if (!inFooterSection)
                    {
                        if (line.Contains("PACKAGE TOTALS   MIX -")) //this is the start of the footer/totals section
                            inFooterSection = true;
                        else
                        {
                            if (!line.StartsWith("              |   |") && !line.StartsWith("---------------"))
                                if (line.Contains("|"))
                                {
                                    List<string> segments = line.Split('|').ToList();

                                    currentPreprintName = segments[0];

                                    Int32 packageCounter = 0;

                                    foreach (string segment in segments.Skip(2))  //skip the first two records
                                    {
                                        if (segment.Trim() == "X")
                                        {
                                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Packages_Preprints",
                                                                        new SqlParameter("@pintPackagesLoadsID", loadsId),
                                                                        new SqlParameter("@pvchrPreprintName", currentPreprintName.Replace("\f", "")),
                                                                        new SqlParameter("@pintPackageNumber", currentPagePackageNumbers[packageCounter]));
                                        }

                                        packageCounter++;
                                    }

                                }
                                else
                                    currentPreprintName = line.Trim();    //this must be an instance where the data record and name are on 2 different lines
                            else if (line.StartsWith("              |   |"))
                            {
                                List<string> segments = line.Replace("              |   |", "").Split('|').ToList();

                                Int32 packageCounter = 0;

                                foreach (string segment in segments)
                                {
                                    if (segment.Trim() == "X")
                                    {
                                        ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Packages_Preprints",
                                                                    new SqlParameter("@pintPackagesLoadsID", loadsId),
                                                                    new SqlParameter("@pvchrPreprintName", currentPreprintName.Replace("\f", "")),
                                                                    new SqlParameter("@pintPackageNumber", currentPagePackageNumbers[packageCounter]));
                                    }

                                    packageCounter++;
                                }
                            }
                        }

                    }
                    else if (inFooterSection) //this is the totals section at the end of the file
                    {
                        if (!line.StartsWith("PKG  QTY"))
                        {
                            List<string> segments = line.Replace("\f", "").Split(' ').ToList();

                            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Update_Packages",
                                                    new SqlParameter("@pintPackagesLoadsID", loadsId),
                                                    new SqlParameter("@pintPackageNumber", segments[0]),
                                                    new SqlParameter("@pintQuantity", segments[2]));
                        }
                    }
                }
            }

            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Update_Packages_Loads",
                                                    new SqlParameter("@pintPackagesLoadsID", loadsId),
                                                    new SqlParameter("@pvchrMixDate", reportDate),
                                                    new SqlParameter("@pvchrMixName", mixName),
                                                    new SqlParameter("@pflgSuccessfullLoad", true));

            WriteToJobLog(JobLogMessageType.INFO, "Load information updated.");


            ExecuteNonQuery(DatabaseConnectionStringNames.Manifests, "Proc_Insert_Packages_Loads_Latest",
                                        new SqlParameter("@pintPackagesLoadsID", loadsId),
                                        new SqlParameter("@psdatMixDate", reportDate),
                                        new SqlParameter("@pvchrMixName", mixName));


        }

    }
}
