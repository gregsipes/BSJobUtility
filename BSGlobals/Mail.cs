using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace BSGlobals
{
   public class Mail
    {
        /// <summary>
        /// Send email. See mail settings in ManagedJobsUtilitySystem section of app.config.
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        /// <param name="bodyIsHTML"></param>
        /// <param name="recipients"></param>
        /// <param name="ccs">Optional</param>
        /// <param name="bccs">Optional</param>
        public static void SendMail(string subject, string body, bool bodyIsHTML, string recipients = null, string ccs = null, string bccs = null, string attachment = null)
        {
            try
            {
                using (SmtpClient client = new SmtpClient())
                {

                    client.Host = Config.GetConfigurationKeyValue("BSJobUtilitySection", "MailHost");

                    using (MailMessage message = new MailMessage())
                    {
                        message.From = new MailAddress(Config.GetConfigurationKeyValue("BSJobUtilitySection", "DefaultSender"));
                        message.Subject = subject;
                        message.Body = body;
                        message.IsBodyHtml = bodyIsHTML;

                        if (attachment != null)
                        {
                            var attach = new Attachment(attachment);
                            message.Attachments.Add(attach);
                        }


                        if (recipients == null)
                            message.To.Add(new MailAddress(Config.GetConfigurationKeyValue("BSJobUtilitySection", "DefaultRecipient")));
                        else
                        {
                            recipients = recipients.Replace(",", ";");

                            foreach (var recipient in recipients.Split(';'))
                            {
                                if (!string.IsNullOrEmpty(recipient))
                                    message.To.Add(new MailAddress(recipient.Trim()));
                            }
                        }

                        if (ccs != null)
                        {
                            ccs = ccs.Replace(",", ";");

                            foreach (var cc in ccs.Split(';'))
                            {
                                if (!string.IsNullOrEmpty(cc))
                                    message.CC.Add(new MailAddress(cc.Trim()));
                            }
                        }

                        if (bccs != null)
                        {

                            bccs = bccs.Replace(",", ";");

                            foreach (var bcc in bccs.Split(';'))
                            {
                                if (!String.IsNullOrEmpty(bcc))
                                    message.Bcc.Add(new MailAddress(bcc.Trim()));
                            }
                        }

                        client.Send(message);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
