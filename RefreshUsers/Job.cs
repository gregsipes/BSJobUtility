using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace RefreshUsers
{
    public class Job : JobBase
    {
        public string Version { get; set; }

        public override void SetupJob()
        {
            JobName = "Refresh Users";
            JobDescription = "TODO";
            AppConfigSectionName = "RefreshUsers";
        }

        public override void ExecuteJob()
        {
            try
            {
                string securityPassPhrase = DeterminePassPhrase(DatabaseConnectionStringNames.ServReq, "ServReqUserSID");
              //  string securityPassPhrase = DeterminePassPhrase(DatabaseConnectionStringNames.Passwords, "PasswordUserSID");
            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }
    }
}
