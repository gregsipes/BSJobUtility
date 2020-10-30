using BSJobBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace DeleteEmptyTMPFiles
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Delete Empty TMP files";
            JobDescription = "Deletes empty temporary files from different locations";
            AppConfigSectionName = "DeleteEmptyTMPFiles";
        }

        public override void ExecuteJob()
        {
            try
            {

                Int32 deletedFileCount = 0;

                List<string> inputDirectories = GetConfigurationKeyValue("InputDirectories").Split(',').ToList();


                foreach (string directory in inputDirectories)
                {
                    List<string> files = Directory.GetFiles(directory, "*.tmp").ToList();

                    //delete each file
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Length == 0)
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"Deleting {file}");
                            File.Delete(file);
                        }
                        deletedFileCount++;
                    }
                }

                if (deletedFileCount > 0)
                    WriteToJobLog(JobLogMessageType.INFO, $"{deletedFileCount} files deleted");

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }
    }
}
