using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace TradeAutoEmail
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Trade Auto Email";
            JobDescription = "Sends out email notifications for requested trade inventory";
            AppConfigSectionName = "TradeAutoEmail";
        }

        public override void ExecuteJob()
        {
            try
            {
                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.Trade, "Proc_Update_Requested_Emails_To_Be_Sent").ToList();

                if (results != null && results.Count() > 0)
                {
                    WriteToJobLog(JobLogMessageType.INFO, "Email requests have been found");

                    foreach (Dictionary<string, object> result in results)
                    {
                        if (!String.IsNullOrEmpty(result["work_email_address"].ToString()))
                        {
                            WriteToJobLog(JobLogMessageType.INFO, $"Sending email to {result["work_email_address"].ToString()}");

                            StringBuilder stringBuilder = new StringBuilder();

                            stringBuilder.AppendLine($"Trade Automatic Email Notification: {result["status_latest"].ToString()}");
                            stringBuilder.AppendLine($"Category: {result["category_description"].ToString()}");
                            stringBuilder.AppendLine($"Subcategory: {result["subcategory_description"].ToString()}");
                            stringBuilder.AppendLine($"Event Date Range: {result["event_date_range"].ToString()}");
                            stringBuilder.AppendLine($"Storage Location: {result["storage_location"].ToString()}");
                            stringBuilder.AppendLine($"Requested By: {result["requested_last_first_department"].ToString()}");
                            stringBuilder.AppendLine($"Requested Date/Time: {result["requested_date_time"].ToString()}");
                            stringBuilder.AppendLine($"Quantity Requested: {result["quantity_requested"].ToString()}");
                            stringBuilder.AppendLine($"Quantity Earmarked: {result["quantity_earmarked"].ToString()}");
                            stringBuilder.AppendLine($"Reserved?: {result["reserved_flag"].ToString()}");

                            SendMail($"Trade Automatic Email Notification: {result["status_latest"].ToString()}", stringBuilder.ToString(), false, result["work_email_address"].ToString());

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
