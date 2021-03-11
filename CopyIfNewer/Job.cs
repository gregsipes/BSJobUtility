using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace CopyIfNewer
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "CopyIfNewer";
            JobDescription = "Checks for and copies over the most the most recent version of files, either file by file or an entire directory";
            AppConfigSectionName = "CopyIfNewer";
        }

        public override void ExecuteJob()
        {
            try
            {
               List<string> files = Directory.GetFiles(GetConfigurationKeyValue("BradburySourceDirectory"), "*").ToList();            
                    
                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                       
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
