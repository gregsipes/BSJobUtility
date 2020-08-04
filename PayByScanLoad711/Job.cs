using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PayByScanLoad711
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Pay By Scan Load - 711";
            JobDescription = "";
            AppConfigSectionName = "PayByScanLoad711";
        }

        public override void ExecuteJob()
        {
            //try
            //{
            //    List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpcarrier.????????").ToList();

            //    files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpcwd.????????").ToList());
            //    files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpfree pub.????????").ToList());
            //    files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpnie.????????").ToList());
            //    files.AddRange(Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "invexpalb1.????????").ToList());


            //    //load configuration from configuration specific tables
            //    Dictionary<string, object> configurationGeneral = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Configuration_General").FirstOrDefault();  //there is only 1 entry in this table
            //    List<Dictionary<string, object>> loadFileConfigurations = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Configuration_Load_Files").ToList();
            //    List<Dictionary<string, object>> configurationTables = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Configuration_Tables").ToList();


            //    //iterate and process files
            //    if (files != null && files.Count() > 0)
            //    {
            //        foreach (string file in files)
            //        {
            //            FileInfo fileInfo = new FileInfo(file);

            //            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "dbo.Proc_Select_Loads_If_Processed",
            //                                                    new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
            //                                                    new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


            //            if (previouslyLoadedFile == null)
            //            {
            //                //make sure we the file is no longer being edited
            //                if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
            //                {
            //                    WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
            //                    CopyAndProcessFile(fileInfo);
            //                }
            //                else
            //                    WriteToJobLog(JobLogMessageType.INFO, "There's a chance the file is still getting updated, so we'll pick it up next run");

            //            }
            //            //else
            //            //{
            //            //    ExecuteNonQuery(DatabaseConnectionStringNames.PBSInvoiceExportLoad, "Proc_Insert_Loads_Not_Loaded",
            //            //                    new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString()),
            //            //                    new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
            //            //                    new SqlParameter("@pdatLastModified", fileInfo.LastWriteTime),
            //            //                    new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
            //            //                    new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
            //            //                    new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString()));
            //            //}
            //        }
            //    }

            //}
            //catch (Exception ex)
            //{
            //    LogException(ex);
            //    throw;
            //}
        }
    }
}
