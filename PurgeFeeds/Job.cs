using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace PurgeFeeds
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "PurgeFeeds";
            JobDescription = "Removes historic Feed build records based on a number of days to retain.";
            AppConfigSectionName = "PurgeFeeds";

        }

        public override void ExecuteJob()
        {
            try
            {
                //get the distinct feeds with their number of days to retain
                List<Dictionary<string, object>> feeds = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Select_Feeds_Distinct_Feed_Types").ToList();

                foreach(Dictionary<string, object> feed in feeds)
                {
                    WriteToJobLog(JobLogMessageType.INFO, "Purging builds for feed type " + feed["feed_type"].ToString() + ", days to keep builds = " + feed["days_to_keep"].ToString());

                    List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Feeds, "Proc_Purge_Builds",
                                                                            new SqlParameter("@pvchrFeedType", feed["feed_type"].ToString()),
                                                                            new SqlParameter("@pintDaysToKeep", feed["days_to_keep"].ToString())).ToList();

                    if (results.Count() <= 0)
                        WriteToJobLog(JobLogMessageType.INFO, "No builds to purge for this feed type");
                    else
                    {
                        foreach(Dictionary<string, object> result in results)
                        {
                            WriteToJobLog(JobLogMessageType.INFO, "Purged " + result[""].ToString());
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
