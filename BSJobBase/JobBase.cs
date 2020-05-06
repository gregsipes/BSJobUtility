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
                Host = GetConfigurationKeyValue("BSJobUtilitySection", "MailHost"),
                DefaultSender = GetConfigurationKeyValue("BSJobUtilitySection", "DefaultSender")
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

        public void LogException(Exception ex)
        {
            SendMail($"Error in Job: {JobName}", ex.ToString(), false);
            WriteToJobLog(JobLogMessageType.ERROR, ex.ToString());
        }

        public void WriteToJobLog(JobLogMessageType type, string message)
        {
            Console.WriteLine($"{DateTime.Now.ToString()} {type.ToString("g"),-7}  Message: {message}");


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
            RunQuery(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
        }

        protected void ExecuteNonQuery(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            RunQuery(connectionStringName, commandType, commandText, parameters);
        }

        private void RunQuery(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            using (SqlCommand command = new SqlCommand())
            {
                try
                {
                    command.Connection = new SqlConnection(GetConnectionStringTo(connectionStringName));
                    command.CommandType = commandType;
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


        #region Excel

        protected void FormatCells(Microsoft.Office.Interop.Excel.Range range, ExcelFormatOption excelFormatOption)
        {
            if (excelFormatOption.StyleName != null)
                range.Style = excelFormatOption.StyleName;
            if (excelFormatOption.NumberFormat != null)
                range.NumberFormat = excelFormatOption.NumberFormat;

            if (excelFormatOption.MergeCells)
                range.Merge();

            range.Font.Bold = excelFormatOption.IsBold;
            range.Font.Underline = excelFormatOption.IsUnderLine;
            range.HorizontalAlignment = excelFormatOption.HorizontalAlignment;


            range.WrapText = excelFormatOption.WrapText;
            range.Interior.Pattern = 1; //solid
            range.Interior.PatternColorIndex = -4105; //automatic
            switch (excelFormatOption.FillColor)
            {
                case ExcelColor.Black:
                    range.Interior.ThemeColor = 2;
                    range.Interior.TintAndShade = 0;
                    break;
                case ExcelColor.LightGray5:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.0499893185216834;
                    break;
                case ExcelColor.LightGray15:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.149998474074526;
                    break;
                case ExcelColor.LightGray25:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.249977111117893;
                    break;
                case ExcelColor.LightGray35:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = -0.349986266670736;
                    break;
                case ExcelColor.LightOrange:
                    range.Interior.ThemeColor = 10;
                    range.Interior.TintAndShade = 0.399975585192419;
                    break;
                case ExcelColor.White:
                    range.Interior.ThemeColor = 1;
                    range.Interior.TintAndShade = 0;
                    break;
                default:
                    range.Interior.ColorIndex = 0;
                    break;
            }

            switch (excelFormatOption.TextColor)
            {
                case ExcelColor.Black:
                    range.Font.ThemeColor = 2;
                    range.Font.TintAndShade = 0;
                    break;
                case ExcelColor.LightGray5:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.0499893185216834;
                    break;
                case ExcelColor.LightGray15:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.149998474074526;
                    break;
                case ExcelColor.LightGray25:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.249977111117893;
                    break;
                case ExcelColor.LightGray35:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = -0.349986266670736;
                    break;
                case ExcelColor.LightOrange:
                    range.Font.ThemeColor = 10;
                    range.Font.TintAndShade = 0.399975585192419;
                    break;
                case ExcelColor.White:
                    range.Font.ThemeColor = 1;
                    range.Font.TintAndShade = 0;
                    break;
            }

            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = excelFormatOption.BorderTopLineStyle;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = excelFormatOption.BorderBottomLineStyle;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = excelFormatOption.BorderLeftLineStyle;
            range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = excelFormatOption.BorderRightLineStyle;

            range.ApplyOutlineStyles();

        }

        protected string ConvertToColumn(Int32 columnNumber)
        {
            Int32 offset = 64;

            if (columnNumber > 256)
                return "";
            else if (columnNumber < 27)
                return ((char)(columnNumber + offset)).ToString();
            else if (columnNumber < 53)
                return "A" + ((char)((columnNumber - 26) + offset)).ToString();
            else if (columnNumber < 79)
                return "B" + ((char)((columnNumber - 52) + offset)).ToString();
            else if (columnNumber < 105)
                return "C" + ((char)((columnNumber - 78) + offset)).ToString();
            else if (columnNumber < 131)
                return "D" + ((char)((columnNumber - 104) + offset)).ToString();
            else if (columnNumber < 157)
                return "E" + ((char)((columnNumber - 130) + offset)).ToString();
            else if (columnNumber < 183)
                return "F" + ((char)((columnNumber - 156) + offset)).ToString();
            else if (columnNumber < 209)
                return "G" + ((char)((columnNumber - 182) + offset)).ToString();
            else if (columnNumber < 235)
                return "H" + ((char)((columnNumber - 208) + offset)).ToString();
            else
                return "I" + ((char)((columnNumber - 234) + offset)).ToString();
        }

        #endregion

        #endregion

        #region Functions

        protected List<Dictionary<string, object>> ExecuteSQL(DatabaseConnectionStringNames connectionStringName, string commandText, params SqlParameter[] parameters)
        {
            return RunSQLCommand(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
        }

        protected List<Dictionary<string, object>> ExecuteSQL(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            return RunSQLCommand(connectionStringName, commandType, commandText, parameters);
        }

        //Change to protected if we ever need to overload. In other words, if we need to pass in something besides a sproc
        private List<Dictionary<string, object>> RunSQLCommand(DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
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
                case DatabaseConnectionStringNames.Wrappers:
                    connectionString = GetConnectionString("wrappers");
                    break;
                case DatabaseConnectionStringNames.Manifests:
                    connectionString = GetConnectionString("manifests");
                    break;
                case DatabaseConnectionStringNames.PBSInvoiceExportLoad:
                    connectionString = GetConnectionString("pbsinvoiceexport");
                    break;
                default:
                    break;
            }

            return connectionString;
        }

       
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
        /// <param name="subject"></param>
        /// <param name="body"></param>
        /// <param name="bodyIsHTML"></param>
        /// <param name="recipients"></param>
        /// <param name="ccs">Optional</param>
        /// <param name="bccs">Optional</param>
        protected void SendMail(string subject, string body, bool bodyIsHTML, string recipients = null, string ccs = null, string bccs = null, string attachment = null)
        {
            try
            {
                using (SmtpClient client = new SmtpClient())
                {

                    client.Host = GetConfigurationKeyValue("BSJobUtilitySection", "MailHost");

                    using (MailMessage message = new MailMessage())
                    {
                        message.From = new MailAddress(GetConfigurationKeyValue("BSJobUtilitySection", "DefaultSender"));
                        message.Subject = subject;
                        message.Body = body;
                        message.IsBodyHtml = bodyIsHTML;

                        if (attachment != null)
                        {
                            var attach = new Attachment(attachment);
                            message.Attachments.Add(attach);
                        }

                        // clean up recipients
                        if (recipients == null)
                            message.To.Add(new MailAddress(GetConfigurationKeyValue("BSJobUtilitySection", "DefaultRecipient")));
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
                            // clean up recipients
                            ccs = ccs.Replace(",", ";");

                            foreach (var cc in ccs.Split(';'))
                            {
                                if (!string.IsNullOrEmpty(cc))
                                    message.CC.Add(new MailAddress(cc.Trim()));
                            }
                        }

                        if (bccs != null)
                        {
                            // clean up recipients
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
