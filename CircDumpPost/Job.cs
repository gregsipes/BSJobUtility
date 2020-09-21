using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace CircDumpPost
{
    public class Job : JobBase
    {
        public int GroupNumber { get; set; }

        public override void SetupJob()
        {
            JobName = "Circ Dump Post";
            JobDescription = "Cleans up any remaning .successful files from the dump directories. This is step 3 in the PBS import process.";
            AppConfigSectionName = "CircDumpPost";
        }

        public override void ExecuteJob()
        {
            try
            {

                //check for any log any unsucessful files
                List<string> files = Directory.GetFiles($"{GetConfigurationKeyValue("GroupTouchDirectoryPath")}\\{GroupNumber.ToString()}\\", "*.unsuccessful").ToList();


                if (files.Count() > 0)
                    WriteToJobLog(JobLogMessageType.INFO, $"No post group routines run since at least one *.unsuccessful file exists in {GetConfigurationKeyValue("GroupTouchDirectoryPath")}\\{GroupNumber.ToString()}\\");
                else
                {

                    //check for any successful files
                    files = Directory.GetFiles($"{GetConfigurationKeyValue("GroupTouchDirectoryPath")}\\{GroupNumber.ToString()}\\", "*.successful").ToList();

                    if (files == null || files.Count() <= 0)
                        WriteToJobLog(JobLogMessageType.INFO, $"No post group routines run since no .successful files exist in {GetConfigurationKeyValue("GroupTouchDirectoryPath")}\\{GroupNumber.ToString()}\\"); //this is a remnant prior to conversion
                    else
                    {
                        //delete the successfully processed touch files
                        foreach (string file in files)
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"Deleting {file}");
                            File.Delete(file);
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

    }
}
