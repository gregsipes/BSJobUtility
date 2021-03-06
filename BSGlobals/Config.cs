﻿using System;
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
                if (section != null && section.AllKeys.Contains(keyName))
                    value = section[keyName].ToString();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to locate key {keyName} in section {sectionName} in configuration file.", ex);
            }

            return value;
        }

        public static bool SetConfigurationKeyValue(string sectionName, string keyName, string value)
        {

            NameValueCollection section = null;

            // Update the local configuration.  If it doesn't already exist this will create
            //   file <application name>.exe.config, into which the updated config parameter(s) will be saved.
            //   The local config values override the global app.config file values.
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
                {
                    var xmlDoc = new System.Xml.XmlDocument();
                    xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
                    string node = @"//configuration/" + sectionName + @"/add[@key='" + keyName + "']";
                    //System.Xml.XmlNode xnode = xmlDoc.SelectSingleNode(node);
                    //System.Xml.XmlAttributeCollection attrs = xnode.Attributes;
                    //attr.Value = value;
                    xmlDoc.SelectSingleNode(node).Attributes["value"].Value = value;
                    xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
                    ConfigurationManager.RefreshSection(sectionName);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to locate section {sectionName} in configuration file.", ex);
            }
            return (true);
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
                case DatabaseConnectionStringNames.BSJobUtility:
                    connectionString = GetConnectionString("bsjobutility");
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
                case DatabaseConnectionStringNames.ManifestsFree:
                    connectionString = GetConnectionString("manifestsfree");
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
                case DatabaseConnectionStringNames.PayByScan:
                    connectionString = GetConnectionString("paybyscan");
                    break;
                case DatabaseConnectionStringNames.PrepackInsertLoad:
                    connectionString = GetConnectionString("prepackinserts");
                    break;
                case DatabaseConnectionStringNames.CircDumpWorkLoad:
                    connectionString = GetConnectionString("circdumpwork_load");
                    break;
                case DatabaseConnectionStringNames.CircDumpWorkPopulate:
                    connectionString = GetConnectionString("circdumpwork_populate");
                    break;
                case DatabaseConnectionStringNames.CircDumpPost:
                    connectionString = GetConnectionString("circdumpwork_post");
                    break;
                case DatabaseConnectionStringNames.PBSDumpAWorkLoad:
                    connectionString = GetConnectionString("pbsdumpawork_load");
                    break;
                case DatabaseConnectionStringNames.PBSDumpAWorkPopulate:
                    connectionString = GetConnectionString("pbsdumpawork_populate");
                    break; 
                case DatabaseConnectionStringNames.PBSDumpPost:
                    connectionString = GetConnectionString("pbsdumpawork_post");
                    break;
                case DatabaseConnectionStringNames.PBSDumpBWork:
                    connectionString = GetConnectionString("pbsdumpbwork");
                    break;
                case DatabaseConnectionStringNames.PBSDumpCWork:
                    connectionString = GetConnectionString("pbsdumpcwork");
                    break;
                case DatabaseConnectionStringNames.Purchasing:
                    connectionString = GetConnectionString("purchasing");
                    break;
                case DatabaseConnectionStringNames.SuppliesWorkLoad:
                    connectionString = GetConnectionString("supplieswork");
                    break;
                case DatabaseConnectionStringNames.PBSDump:
                    connectionString = GetConnectionString("pbsdump");
                    break;
                case DatabaseConnectionStringNames.BNTransactions:
                    connectionString = GetConnectionString("bntransactions");
                    break;
                case DatabaseConnectionStringNames.TradeWorkLoad:
                    connectionString = GetConnectionString("tradework");
                    break;
                case DatabaseConnectionStringNames.SubBalanceLoad:
                    connectionString = GetConnectionString("subbalancework");
                    break;
                case DatabaseConnectionStringNames.Feeds:
                    connectionString = GetConnectionString("feeds");
                    break;
                case DatabaseConnectionStringNames.ISInventory:
                    connectionString = GetConnectionString("isinventory");
                    break;
                case DatabaseConnectionStringNames.SynergyReportMaintenance:
                    connectionString = GetConnectionString("synergyreportmaintenance");
                    break;
                case DatabaseConnectionStringNames.SBSJournalEntryImport:
                    connectionString = GetConnectionString("sbsjournalentryimport");
                    break;
                case DatabaseConnectionStringNames.AutoPrintPDF:
                    connectionString = GetConnectionString("autoprintpdf");
                    break;
                case DatabaseConnectionStringNames.PBSGeneral:
                    connectionString = GetConnectionString("pbsgeneral");
                    break;
                case DatabaseConnectionStringNames.BSConsole:
                    connectionString = GetConnectionString("bsconsole");
                    break;
                case DatabaseConnectionStringNames.Brainworks_Work:
                    connectionString = GetConnectionString("brainworks_work");
                    break;
                case DatabaseConnectionStringNames.MailTops:
                    connectionString = GetConnectionString("mailtops");
                    break;
                case DatabaseConnectionStringNames.Newshole:
                    connectionString = GetConnectionString("newshole");
                    break;
                case DatabaseConnectionStringNames.OfficePay_Archived:
                    connectionString = GetConnectionString("officepayarchived");
                    break;
                case DatabaseConnectionStringNames.Palletizers:
                    connectionString = GetConnectionString("palletizers");
                    break;
                case DatabaseConnectionStringNames.PBSInvoiceExport:
                    connectionString = GetConnectionString("pbsinvoiceexport");
                    break;
                case DatabaseConnectionStringNames.Postings:
                    connectionString = GetConnectionString("postings");
                    break;
                case DatabaseConnectionStringNames.Preprints:
                    connectionString = GetConnectionString("preprints");
                    break;
                case DatabaseConnectionStringNames.TouchControl:
                    connectionString = GetConnectionString("touchcontrol");
                    break;
                case DatabaseConnectionStringNames.Trade:
                    connectionString = GetConnectionString("trade");
                    break;
                case DatabaseConnectionStringNames.Passwords:
                    connectionString = GetConnectionString("passwords");
                    break;
                case DatabaseConnectionStringNames.ServReq:
                    connectionString = GetConnectionString("servreq");
                    break;
                default:
                    break;
            }

            return connectionString;
        }

    }
}
