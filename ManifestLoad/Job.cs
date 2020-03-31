using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManifestLoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Manifest Load";
            JobDescription = "Builds manifest labels";
            AppConfigSectionName = "ManifestLoad";
        }

        public override void ExecuteJob()
        {
            try
            {

            }
            catch (Exception ex)
            {
                SendMail($"Error in Job: {JobName}", ex.ToString(), false);
                throw;
            }
        }


    }
}
