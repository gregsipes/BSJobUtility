using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PurgeFeeds
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "PurgeFeeds";
            JobDescription = "Cleans up feeds records";
            AppConfigSectionName = "PurgeFeeds";

        }

        public override void ExecuteJob()
        {
            try
            {
                
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }
    }
}
