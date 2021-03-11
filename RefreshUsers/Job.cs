using BSJobBase;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;
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
                string securityPassPhrase = "";
                DatabaseConnectionStringNames connectionString;

                if (Version == "Passwords") {
                    connectionString = DatabaseConnectionStringNames.Passwords;
                    securityPassPhrase = DeterminePassPhrase(connectionString, "PasswordUserSID");
                }
                else {
                    connectionString = DatabaseConnectionStringNames.ServReq;
                    securityPassPhrase = DeterminePassPhrase(connectionString, "ServReqUserSID");

                }

                WriteToJobLog(JobLogMessageType.INFO, "Clearing refreshusers_active_user_flag");
                ExecuteNonQuery(connectionString, "Proc_Update_Users_RefreshUsers_Active_User_Flag");

                List<string> domainControllers = GetConfigurationKeyValue("DomainControllers").Split(',').ToList();

                foreach (string domainController in domainControllers)
                {
                    string filter = "(cn=BSOU*)";

                    Console.WriteLine(filter);

                    DirectorySearcher searcher = new DirectorySearcher(filter);

                    StringBuilder stringBuilder = new StringBuilder();
                    foreach (SearchResult result in searcher.FindAll())
                    {
                        var userEntry = result.GetDirectoryEntry();
                        stringBuilder.AppendLine(userEntry.Properties["SAMAccountName"].Value.ToString());
                        
                        Console.WriteLine(userEntry.Properties["SAMAccountName"].Value.ToString());
                    }

                    File.WriteAllText("C:\\temp\\test.txt", stringBuilder.ToString());

                }


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                LogException(ex);
                throw;
            }
        }
    }
}
