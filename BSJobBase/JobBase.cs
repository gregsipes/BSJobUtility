using BSJobBase.Classes;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BSJobBase
{
    /// <summary>
    /// This is the top-level base class that all jobs will inherit from
    /// </summary>
    public abstract class JobBase
    {
        #region Properties

        /// <summary>
        /// Holds the unique job name
        /// </summary>
        protected string JobName { get; set; }
        /// <summary>
        /// Holds the full job description
        /// </summary>
        protected string JobDescription { get; set; }
        /// <summary>
        /// Holds the values of the config section in the app.config file
        /// </summary>
        protected string AppConfigSectionName { get; set; }
        /// <summary>
        /// Holds an array of arguments passed into the executable
        /// </summary>
        protected string[] Args { get; set; }
        /// <summary>
        /// Holds the class containing the smtp mail settings
        /// </summary>
        public MailSettings MailSettings { get; set; }
        /// <summary>
        /// Holds the class containing all general settings
        /// </summary>
        public GeneralSettings GeneralSettings { get; set; }
        /// <summary>
        /// Holds the class containing FTP/SFTP settings
        /// </summary>
        public FTPSettings FtpSettings { get; set; }

        #endregion

        #region Constructor

        public JobBase()
        {
            //setup base job settings
            MailSettings = new MailSettings()
            {
                UseTLS = !string.IsNullOrEmpty(GetConfigurationKeyValue("BSJobUtilitySection", "UseTLS")),
                Host = GetConfigurationKeyValue("BSJobUtilitySection", "MailHost"),
                Port = int.Parse(GetConfigurationKeyValue("BSJobUtilitySection", "MailPort")),
                User = GetConfigurationKeyValue("BSJobUtilitySection", "MailUser"),
                Password = GetConfigurationKeyValue("BSJobUtilitySection", "MailPassword"),
                DefaultSender = GetConfigurationKeyValue("BSJobUtilitySection", "DefaultSender"),
                DefaultRecipient = GetConfigurationKeyValue("BSJobUtilitySection", "DefaultRecipient"),
            };
            GeneralSettings = new GeneralSettings()
            {
                DefaultSQLCommandTimeout = int.Parse(GetConfigurationKeyValue("BSJobUtilitySection", "DefaultSQLCommandTimeout"))
            };
            FtpSettings = new FTPSettings()
            {
                ServerResponseTimeoutInSeconds = int.Parse(GetConfigurationKeyValue("BSJobUtilitySection", "FtpSftpServerResponseTimeoutInSeconds"))
            };

        }

        #endregion

        #region Abstract

        public abstract void SetupJob();

        public abstract void ExecuteJob();


        #endregion

        #region Virtual Methods

        public virtual void PreExecuteJob(string[] args)
        {
            // validate that job has been setup correctly
            if (JobName == null || JobName == "")
                throw new Exception("Job Name property can not be empty.");

            if (JobDescription == null || JobDescription == "")
                throw new Exception("Job Description property can not be empty.");

            if (AppConfigSectionName == null || AppConfigSectionName == "")
                throw new Exception("App Config Section Name property can not be empty.");

            // store passed in commandline args
            Args = args;

            // setup log expiration for purge
            int logExpiration = 27;

            // if job has a custom log entry age limit then read it and store it in the database
            // log expiration age limit is used by the log purge job
            // optional so check if it exists before using it
            try
            {
                string tempValue = GetConfigurationKeyValue("LogEntryAgeLimitWeeks");

                if (tempValue != null)
                    logExpiration = int.Parse(tempValue);
            }
            catch (Exception)
            {
                // eat exceptions here
            }

            // basic logging
            WriteToJobLog(JobLogMessageType.INFO, "Job starting");
        }

        public virtual void PostExecuteJob()
        {
            WriteToJobLog(JobLogMessageType.INFO, "Job completed");
        }

        #endregion

        #region Methods

        public void WriteToJobLog(JobLogMessageType type, string message)
        {
            Console.WriteLine($"[Log] Type: {type.ToString("g"),-7}  Message: {message}");


            using (SqlCommand command = new SqlCommand())
            {
                try
                {
                    command.Connection = new SqlConnection(GetConnectionStringTo(DatabaseConnectionStringNames.EventLogs));
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "dbo.InsertJobLog";
                    command.Parameters.Add(new SqlParameter("@JobName", JobName));
                    command.Parameters.Add(new SqlParameter("@MessageType", type.ToString("d")));
                    command.Parameters.Add(new SqlParameter("@Message", message));

                    command.Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error inserting log record. {ex.Message}");
                }
                finally
                {
                    if (command != null && command.Connection != null)
                        command.Connection.Close();
                }
            }
        }

        protected void ExecuteNonQuery(DatabaseConnectionStringNames connectionStringName, string commandText, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                try
                {
                    command.Connection = new SqlConnection(GetConnectionStringTo(connectionStringName));
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = commandText;
                    command.CommandTimeout = 0;

                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            command.Parameters.Add(param);
                        }
                    }

                    command.Connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error executing query. {ex.Message}");
                }
                finally
                {
                    if (command != null && command.Connection != null)
                        command.Connection.Close();
                }
            }
        }

        #endregion

        #region Functions

        protected List<Dictionary<string, object>> ExecuteSQL(DatabaseConnectionStringNames connectionStringName, string commandText, params SqlParameter[] parameters)
        {
            return ExecuteSQL(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
        }

        //Change to protected if we ever need to overload. In other words, if we need to pass in something besides a sproc
        private List<Dictionary<string, object>> ExecuteSQL(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {

            List<Dictionary<string, object>> rowsToReturn = new List<Dictionary<string, object>>();

            using (SqlDataReader reader = ExecuteQuery(connectionStringName, commandType, commandText, parameters))
            {
                while (reader.Read())
                {
                    Dictionary<string, object> dictionary = new Dictionary<string, object>();

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        dictionary.Add(reader.GetName(i), reader.GetValue(i));
                    }

                    rowsToReturn.Add(dictionary);
                }
            }

            return rowsToReturn;
        }

        private SqlDataReader ExecuteQuery(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                command.Connection = new SqlConnection(GetConnectionStringTo(connectionStringName));
                command.CommandType = commandType;
                command.CommandText = commandText;

                if (parameters != null)
                {
                    foreach (var param in parameters)
                    {
                        command.Parameters.Add(param);  //new SqlParameter(param.Key, param.Value)
                    }
                }
                command.Connection.Open();

                //https://docs.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqlcommand?redirectedfrom=MSDN&view=netframework-4.6
                // When using CommandBehavior.CloseConnection, the connection will be closed when the 
                // IDataReader is closed.
                SqlDataReader reader = command.ExecuteReader(CommandBehavior.CloseConnection);

                return reader;

            }
        }

        /// <summary>
        /// Returns the value of key from the app.config file, uses AppConfigSectionName property for section name
        /// </summary>
        /// <param name="keyName"></param>
        /// <returns></returns>
        protected string GetConfigurationKeyValue(string keyName)
        {
            return GetConfigurationKeyValue(AppConfigSectionName, keyName);
        }

        /// <summary>
        /// Returns the value of key from the app.config file
        /// </summary>
        /// <param name="sectionName"></param>
        /// <param name="keyName"></param>
        /// <returns></returns>
        protected string GetConfigurationKeyValue(string sectionName, string keyName)
        {
            NameValueCollection section = null;
            string value = null;

            try
            {
                section = ConfigurationManager.GetSection(sectionName) as NameValueCollection;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to locate section {sectionName} in configuration file.", ex);
            }

            try
            {
                if (section != null)
                    value = section[keyName].ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to locate key {keyName} in section {sectionName} in configuration file.", ex);
            }

            return value;
        }

        /// <summary>
        /// Returns connection string from the app.config file
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string GetConnectionString(string name)
        {
            return ConfigurationManager.ConnectionStrings[name].ConnectionString;
        }

        /// <summary>
        /// Returns connection string from the app.config file
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetConnectionStringTo(DatabaseConnectionStringNames name)
        {
            string connectionString = null;

            switch (name)
            {
                case DatabaseConnectionStringNames.EventLogs:
                    connectionString = GetConnectionString("eventlogs");
                    break;
                case DatabaseConnectionStringNames.Parking:
                    connectionString = GetConnectionString("parking");
                    break;
                case DatabaseConnectionStringNames.SBSReports:
                    connectionString = GetConnectionString("sbsreports");
                    break;
                case DatabaseConnectionStringNames.PBS2Macro:
                    connectionString = GetConnectionString("pbs2macro");
                    break;
                case DatabaseConnectionStringNames.Commissions:
                    connectionString = GetConnectionString("commissions");
                    break;
                case DatabaseConnectionStringNames.BuffNewsForBW:
                    connectionString = GetConnectionString("buffnewsforbw");
                    break;
                case DatabaseConnectionStringNames.Brainworks:
                    connectionString = GetConnectionString("brainworks");
                    break;
                case DatabaseConnectionStringNames.BARC:
                    connectionString = GetConnectionString("barc");
                    break;
                case DatabaseConnectionStringNames.CommissionsRelated:
                    connectionString = GetConnectionString("commissionsrelated");
                    break;
                default:
                    break;
            }

            return connectionString;
        }

        //public static bool IsValidFileNameAndPath(string filePath)
        //{
        //    if (!IsStringValidDirectory(filePath))
        //    {
        //        return false;
        //    }

        //    var invalidChars = Path.GetInvalidFileNameChars();

        //    foreach (var c in invalidChars)
        //    {
        //        if (Path.GetFileName(filePath).Contains(c))
        //        {
        //            return false;
        //        }
        //    }

        //    return true;
        //}

        //public static bool IsStringValidDirectory(string filePath)
        //{
        //    if (string.IsNullOrWhiteSpace(filePath))
        //        return false;

        //    //check for invalid path characters
        //    if (filePath.ToCharArray().Where(c => Path.GetInvalidPathChars().Contains(c)).Count() > 0)
        //    {
        //        return false;
        //    }

        //    try
        //    {
        //        //make sure we can extract a directory
        //        Path.GetDirectoryName(filePath);
        //    }
        //    catch (Exception)
        //    {
        //        return false;
        //    }

        //    //We've passed all tests, the path is valid
        //    return true;
        //}

        public static void CheckCreateDirectory(string filePath)
        {
            CheckCreateDirectory(filePath, false);
        }

        public static void CheckCreateDirectory(string filePath, bool containsFileName)
        {
            string directory = "";
            if (containsFileName)
                directory = Path.GetDirectoryName(filePath);
            else
                directory = filePath;


            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }

        /// <summary>
        /// Validates a string of email addresses by attempting to create a MailAddress object for each item.
        /// </summary>
        /// <param name="entry"></param>
        /// <returns>True if email addresses contained in string are all valid.</returns>
        //public static bool IsValidEmailAddressOrChainOfEmailAddresses(string entry)
        //{
        //    return IsValidEmailAddressOrChainOfEmailAddresses(entry, ';');
        //}

        /// <summary>
        /// Validates a string of email addresses by attempting to create a MailAddress object for each item.
        /// </summary>
        /// <param name="entry"></param>
        /// <param name="delimiter"></param>
        /// <returns>True if email addresses contained in string are all valid.</returns>
        //public static bool IsValidEmailAddressOrChainOfEmailAddresses(string entry, char delimiter)
        //{
        //    if (string.IsNullOrWhiteSpace(entry))
        //        return false;

        //    // catch extra delimiter at end of string
        //    if (entry.Trim().Last() == delimiter)
        //        entry = entry.Trim().Substring(0, entry.Trim().Length - 1);

        //    foreach (var item in entry.Split(delimiter))
        //    {
        //        try
        //        {
        //            var address = new MailAddress(item);
        //        }
        //        catch (Exception)
        //        {
        //            return false;
        //        }
        //    }

        //    return true;
        //}

        //internal static void CheckDeleteFile(string file)
        //{
        //    if (File.Exists(file))
        //    {
        //        File.Delete(file);
        //    }
        //}

        //internal static void CheckCreateFile(string fileNameAndPath)
        //{
        //    if (!File.Exists(fileNameAndPath))
        //    {
        //        using (var stream = File.Create(fileNameAndPath)) { }
        //    }
        //}

        //public static List<string> GetFiles(string sourceDirectory)
        //{
        //    // validate existence of directory
        //    CheckCreateDirectory(sourceDirectory);

        //    return Directory.GetFiles(sourceDirectory)
        //        .ToList();
        //}

        public static List<string> GetFiles(string sourceDirectory, Regex reg)
        {
            // validate existence of directory
            CheckCreateDirectory(sourceDirectory);

            return Directory.GetFiles(sourceDirectory)
                .Where(f => ((reg == null) ? true : reg.IsMatch(Path.GetFileName(f))))
                .ToList();
        }

        //public static List<string> GetFiles(string sourceDirectory, int ageLimitMinutes)
        //{
        //    // validate existence of directory
        //    CheckCreateDirectory(sourceDirectory);

        //    return Directory.GetFiles(sourceDirectory)
        //        .Where(f => File.GetCreationTime(f) < DateTime.Now.AddMinutes(ageLimitMinutes * -1))
        //        .ToList();
        //}

        //public static List<string> GetFiles(string sourceDirectory, Regex reg, int ageLimitMinutes)
        //{
        //    // validate existence of directory
        //    CheckCreateDirectory(sourceDirectory);

        //    return Directory.GetFiles(sourceDirectory)
        //        .Where(f => ((reg == null) ? true : reg.IsMatch(Path.GetFileName(f))) && File.GetCreationTime(f) < DateTime.Now.AddMinutes(ageLimitMinutes * -1))
        //        .ToList();
        //}

        /// <summary>
        /// Send email using default sender email address. See DefaultSender in ManagedJobsUtilitySystem section of app.config.
        /// </summary>
        /// <param name="recipients"></param>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        /// <param name="bodyIsHTML"></param>
        /// <param name="ccs">Optional</param>
        /// <param name="bccs">Optional</param>
        //protected void SendMail(string recipients, string subject, string body, bool bodyIsHTML, string ccs = null, string bccs = null, string attachment = null)
        //{
        //    string from = GetConfigurationKeyValue("BSJobUtility", "DefaultSender");
        //    SendMail(from, recipients, subject, body, bodyIsHTML, ccs, bccs, attachment);
        //}

        /// <summary>
        /// Send email. See mail settings in ManagedJobsUtilitySystem section of app.config.
        /// </summary>
        /// <param name="from"></param>
        /// <param name="recipients"></param>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        /// <param name="bodyIsHTML"></param>
        /// <param name="ccs">Optional</param>
        /// <param name="bccs">Optional</param>
        //protected void SendMail(string from, string recipients, string subject, string body, bool bodyIsHTML, string ccs = null, string bccs = null, string attachment = null)
        //{
        //    try
        //    {
        //        using (SmtpClient client = new SmtpClient())
        //        {
        //            NetworkCredential creds = new NetworkCredential
        //            {
        //                UserName = GetConfigurationKeyValue("ManagedJobsUtilitySystem", "MailUser"),
        //                Password = GetConfigurationKeyValue("ManagedJobsUtilitySystem", "MailPassword")
        //            };

        //            client.Host = GetConfigurationKeyValue("ManagedJobsUtilitySystem", "MailHost");
        //            client.Port = int.Parse(GetConfigurationKeyValue("ManagedJobsUtilitySystem", "MailPort"));
        //            client.Credentials = creds;
        //            client.EnableSsl = !string.IsNullOrWhiteSpace(GetConfigurationKeyValue("ManagedJobsUtilitySystem", "UseTLS"));

        //            using (MailMessage message = new MailMessage())
        //            {
        //                message.From = new MailAddress(from);
        //                message.Subject = subject;
        //                message.Body = body;
        //                message.IsBodyHtml = bodyIsHTML;

        //                if (attachment != null)
        //                {
        //                    var attach = new Attachment(attachment);
        //                    message.Attachments.Add(attach);
        //                }

        //                // clean up recipients
        //                recipients = recipients.Replace(",", ";");

        //                foreach (var recipient in recipients.Split(';'))
        //                {
        //                    if (!string.IsNullOrEmpty(recipient))
        //                        message.To.Add(new MailAddress(recipient.Trim()));
        //                }

        //                if (ccs != null)
        //                {
        //                    // clean up recipients
        //                    ccs = ccs.Replace(",", ";");

        //                    foreach (var cc in ccs.Split(';'))
        //                    {
        //                        if (!string.IsNullOrEmpty(cc))
        //                            message.CC.Add(new MailAddress(cc.Trim()));
        //                    }
        //                }

        //                if (bccs != null)
        //                {
        //                    // clean up recipients
        //                    bccs = bccs.Replace(",", ";");

        //                    foreach (var bcc in bccs.Split(';'))
        //                    {
        //                        if (!String.IsNullOrEmpty(bcc))
        //                            message.Bcc.Add(new MailAddress(bcc.Trim()));
        //                    }
        //                }

        //                client.Send(message);
        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //}

        //protected string GenerateEmailBodyAsHTML(string bodyText)
        //{
        //    return GenerateEmailBodyAsHTML(null, bodyText, null);
        //}

        //protected string GenerateEmailBodyAsHTML(string headerText, string bodyText, string footerText)
        //{
        //    StringBuilder builder = new StringBuilder();

        //    try
        //    {
        //        builder.Append("<meta http-equiv=\"Content - Type\" content=\"text / html; charset = iso - 8859 - 1\">");

        //        if (headerText != null && headerText != "")
        //        {
        //            builder.Append(headerText.Replace(Environment.NewLine, "<br>"));
        //            builder.Append("<br>");
        //        }

        //        if (bodyText != null && bodyText != "")
        //        {
        //            builder.Append(bodyText.Replace(Environment.NewLine, "<br>"));
        //            builder.Append("<br>");
        //        }

        //        if (footerText != null && footerText != "")
        //        {
        //            builder.Append(footerText.Replace(Environment.NewLine, "<br>"));
        //            builder.Append("<br>");
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }

        //    return builder.ToString();
        //}

        //public int DefaultSQLCommandTimeout
        //{
        //    get
        //    {
        //        int returnValue = 60;

        //        try
        //        {
        //            returnValue = int.Parse(GetConfigurationKeyValue("BSJobUtilitySection", "DefaultSQLCommandTimeout"));
        //        }
        //        catch (Exception)
        //        {

        //            throw;
        //        }

        //        return returnValue;
        //    }
        //}

        //protected string GenerateEmailBodyAsHTML(string bodyHeaderText, List<EmailTableHeader> tableHeaders, IEnumerable<EmailTableBodyItemBase> tableRows)
        //{
        //    return GenerateEmailBodyAsHTML(bodyHeaderText, tableHeaders, tableRows, "");
        //}

        //protected string GenerateEmailBodyAsHTML(string bodyHeaderText, List<EmailTableHeader> tableHeaders, IEnumerable<EmailTableBodyItemBase> tableRows, string bodyFooterText)
        //{
        //    StringBuilder builder = new StringBuilder();

        //    try
        //    {
        //        builder.Append("<meta http-equiv=\"Content - Type\" content=\"text / html; charset = iso - 8859 - 1\">");
        //        builder.Append(bodyHeaderText == null ? "" : bodyHeaderText);
        //        builder.Append("<br>");
        //        builder.Append(GenerateTableAsHTML(tableHeaders, tableRows));
        //        builder.Append("<br>");
        //        builder.Append(bodyFooterText == null ? "" : bodyFooterText);
        //        builder.Append("<br>");
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }

        //    return builder.ToString();
        //}

        //protected string GenerateTableAsHTML(List<EmailTableHeader> tableHeaders, IEnumerable<EmailTableBodyItemBase> tableRows)
        //{
        //    StringBuilder builder = new StringBuilder();

        //    if (tableHeaders != null && tableRows != null)
        //    {
        //        double columnWidth = 1 / tableHeaders.Count;

        //        builder.Append("<table border=\"3\" style=\"border - collapse: collapse\">");

        //        // write table header row
        //        builder.Append("<tr>");

        //        foreach (EmailTableHeader item in tableHeaders)
        //        {
        //            builder.Append(string.Format("<td width=\"{0}%\">", columnWidth));
        //            builder.Append(string.Format("<b>{0}</b>", item.HeaderText));
        //            builder.Append("</td>");
        //        }

        //        builder.Append("</tr>");

        //        string backgroundColor = "White";

        //        // write table body rows
        //        foreach (EmailTableBodyItemBase item in tableRows)
        //        {
        //            builder.Append("<tr>");

        //            foreach (EmailTableBodyCell subItem in item.ItemsToListOfEmailTableBodyCells())
        //            {
        //                switch (subItem.Type)
        //                {
        //                    case EmailCellType.NONE:
        //                        backgroundColor = "White";
        //                        break;
        //                    case EmailCellType.INFO:
        //                        backgroundColor = "Silver";
        //                        break;
        //                    case EmailCellType.ERROR:
        //                        backgroundColor = "Red";
        //                        break;
        //                    default:
        //                        break;
        //                }

        //                builder.Append($"<td colspan=\"{subItem.ColumnSpan}\" style=\"background-color: {backgroundColor};\">");

        //                if (subItem.ContentIsHTML)
        //                    builder.Append($"<b>{subItem.Content}</b>");
        //                else
        //                    builder.Append($"<b>{WebUtility.HtmlEncode(subItem.Content)}</b>");

        //                builder.Append("</td>");
        //            }

        //            builder.Append("</tr>");
        //        }

        //        builder.Append("</table>");
        //    }

        //    return builder.ToString();
        //}


        #endregion


    }


}
