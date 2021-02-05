using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;
using System.Data.SqlClient;
using System.IO;
using System.Xml.Linq;

namespace SBSReportsLoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "SBSReportsLoad";
            JobDescription = "TODO";
            AppConfigSectionName = "SBSReportsLoad";

        }

        public override void ExecuteJob()
        {
            try
            {
                //string securityPassPhrase = DeterminePassPhrase(DatabaseConnectionStringNames.SBSReports);

                ////throw an exception if the passphrase comes back empty or null. This is used later to decrypt 
                //if (String.IsNullOrEmpty(securityPassPhrase))
                //    throw new Exception($"Invalid passphrase for user {System.Security.Principal.WindowsIdentity.GetCurrent().Name}");

                //create a load record and return the unique id
                Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Loads").FirstOrDefault();
                Int64 loadsId = Convert.ToInt64(result["loads_id"].ToString());
                WriteToJobLog(JobLogMessageType.INFO, $"Loads Id: {loadsId}");

                //get all unqiue table names
                List<Dictionary<string, object>> tables = ExecuteSQL(DatabaseConnectionStringNames.SBSReports, "Proc_Select_Dictionary_Unique_Table_Names",
                                                                            new SqlParameter("@pvchrTableName", "")).ToList();

                foreach (Dictionary<string, object> table in tables)
                {
                    WriteToJobLog(JobLogMessageType.INFO, $"Processing {table["table_name"].ToString()}");

                    WriteToJobLog(JobLogMessageType.INFO, "Retrieving column names");
                    List<Dictionary<string, object>> fields = ExecuteSQL(DatabaseConnectionStringNames.SBSReports, "Proc_Select_Dictionary_For_Table_Name",
                                                                                        new SqlParameter("@pvchrTableName", table["table_name"].ToString())).ToList();

                    //get XML file
                    string xmlFile = GetConfigurationKeyValue("InputDirectory") + table["table_name"].ToString() + ".xml";

                    if (File.Exists(xmlFile))
                    {
                        XDocument xml = XDocument.Load(xmlFile);

                        List<XElement> nodes = new List<XElement>(); // xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").ToList();


                        //this case statement replaces the replaces the where_conditions table. We ran into issues converting the sql strings into the Linq To XML queries,
                        //so for the sake of time, we moved the where clauses here
                        switch (table["table_name"].ToString().ToLower())
                        {
                            case "empded2":
                                nodes = xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").Where(n => n.Elements("DeductCode").ToString().Contains("uf")).ToList();
                                break;
                            case "employee":
                                nodes = xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").Where(n => n.Elements("CompanyId") != null && n.Elements("CompanyId").ToString().Contains("BNEWS")).ToList();
                                break;
                            case "tcard2":
                                xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").Where(n => Convert.ToDateTime(n.Elements("TrxDate").ToString()) >= DateTime.Now.AddYears(-3)).ToList();
                                break;
                            default:
                                nodes = xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").ToList();
                                break;
                        }

                        foreach (XElement node in nodes)
                        {


                            foreach (Dictionary<string, object> field in fields)
                            {

                            }
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
