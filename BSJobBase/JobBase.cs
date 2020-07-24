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
            WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, "Job starting");
        }

        public virtual void PostExecuteJob()
        {
            WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, "Job completed");
        }

        #endregion

        #region Methods

        public void LogException(Exception ex)
        {
            BSGlobals.Mail.SendMail($"Error in Job: {JobName}", ex.ToString(), false);
            BSGlobals.DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.ERROR, ex.ToString(), JobName);
        }



        public void WriteToJobLog(BSGlobals.Enums.JobLogMessageType type, string message)
        {
            BSGlobals.DataIO.WriteToJobLog(type, message, JobName);           
        }

        protected void ExecuteNonQuery(BSGlobals.Enums.DatabaseConnectionStringNames connectionStringName, string commandText, params SqlParameter[] parameters)
        {
            BSGlobals.DataIO.ExecuteNonQuery(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
        }

        protected void ExecuteNonQuery(BSGlobals.Enums.DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            BSGlobals.DataIO.ExecuteNonQuery(connectionStringName, commandType, commandText, parameters);
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

        protected List<Dictionary<string, object>> ExecuteSQL(BSGlobals.Enums.DatabaseConnectionStringNames connectionStringName, string commandText, params SqlParameter[] parameters)
        {
            // return RunSQLCommand(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
            return BSGlobals.DataIO.ExecuteSQL(connectionStringName, CommandType.StoredProcedure, commandText, parameters);
        }

        protected List<Dictionary<string, object>> ExecuteSQL(BSGlobals.Enums.DatabaseConnectionStringNames connectionStringName, CommandType commandType, string commandText, params SqlParameter[] parameters)
        {
            // return RunSQLCommand(connectionStringName, commandType, commandText, parameters);
            return BSGlobals.DataIO.ExecuteSQL(connectionStringName, commandType, commandText, parameters);
        }

        /// <summary>
        /// Returns the value of key from the app.config file, uses AppConfigSectionName property for section name
        /// </summary>
        /// <param name="keyName"></param>
        /// <returns></returns>
        protected string GetConfigurationKeyValue(string keyName)
        {
            return BSGlobals.Config.GetConfigurationKeyValue(AppConfigSectionName, keyName);
        }

        /// <summary>
        /// Returns the value of key from the app.config file
        /// </summary>
        /// <param name="sectionName"></param>
        /// <param name="keyName"></param>
        /// <returns></returns>
        protected string GetConfigurationKeyValue(string sectionName, string keyName)
        {
            return BSGlobals.Config.GetConfigurationKeyValue(sectionName, keyName);
        }

        /// <summary>
        /// Returns connection string from the app.config file
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string GetConnectionString(string name)
        {
            return BSGlobals.Config.GetConnectionString(name);
        }

        /// <summary>
        /// Returns connection string from the app.config file
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static string GetConnectionStringTo(BSGlobals.Enums.DatabaseConnectionStringNames name)
        {
            return BSGlobals.Config.GetConnectionStringTo(name);            
        }


        public static void CheckCreateDirectory(string filePath)
        {
            BSGlobals.FileIO.CheckCreateDirectory(filePath, false);
        }

        public static void CheckCreateDirectory(string filePath, bool containsFileName)
        {
            BSGlobals.FileIO.CheckCreateDirectory(filePath, containsFileName);
        }

        public static List<string> GetFiles(string sourceDirectory, Regex reg)
        {
           return BSGlobals.FileIO.GetFiles(sourceDirectory, reg);
        }


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
            BSGlobals.Mail.SendMail(subject, body, bodyIsHTML, recipients, ccs, bccs, attachment);            
        }

        protected object FormatNumber(string inputString)
        {
            return BSGlobals.Types.FormatNumber(inputString);
        }

        protected object FormatDateTime(string inputString)
        {
            return BSGlobals.Types.FormatDateTime(inputString);
        }

        protected object FormatString(string inputString)
        {
            return BSGlobals.Types.FormatString(inputString);
        }

        #endregion


    }


}
