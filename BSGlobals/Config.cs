using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace BSGlobals
{
   public  class Config
    {
        /// <summary>
        /// Returns the value of key from the app.config file
        /// </summary>
        /// <param name="sectionName"></param>
        /// <param name="keyName"></param>
        /// <returns></returns>
        public static string GetConfigurationKeyValue(string sectionName, string keyName)
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
        public static string GetConnectionString(string name)
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
                case DatabaseConnectionStringNames.QualificationReportLoad:
                    connectionString = GetConnectionString("qualificationreport");
                    break;
                case DatabaseConnectionStringNames.OfficePay:
                    connectionString = GetConnectionString("officepay");
                    break;
                case DatabaseConnectionStringNames.AutoRenew:
                    connectionString = GetConnectionString("autorenew");
                    break;
                case DatabaseConnectionStringNames.PressRoom:
                    connectionString = GetConnectionString("pressroom");
                    break;
                case DatabaseConnectionStringNames.PressRoomFree:
                    connectionString = GetConnectionString("pressroomfree");
                    break;
                case DatabaseConnectionStringNames.PBSInvoiceTotals:
                    connectionString = GetConnectionString("pbsinvoicetotals");
                    break;
                case DatabaseConnectionStringNames.PBSInvoices:
                    connectionString = GetConnectionString("pbsinvoices");
                    break;
                case DatabaseConnectionStringNames.DMMail:
                    connectionString = GetConnectionString("dmmail");
                    break;
                case DatabaseConnectionStringNames.DMMailFree:
                    connectionString = GetConnectionString("dmmailfree");
                    break;
                default:
                    break;
            }

            return connectionString;
        }
    }
}
