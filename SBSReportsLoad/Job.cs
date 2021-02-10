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

                    //WriteToJobLog(JobLogMessageType.INFO, "Retrieving column names");
                    //List<Dictionary<string, object>> fields = ExecuteSQL(DatabaseConnectionStringNames.SBSReports, "Proc_Select_Dictionary_For_Table_Name",
                    //                                                                    new SqlParameter("@pvchrTableName", table["table_name"].ToString())).ToList();

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
                            //it looks like there are currently 2 encrypted fields, but they are both empty, perhaps this was all disabled at one point?
                            // List<Dictionary<string, object>> encryptedFields = fields.Where(f => Convert.ToBoolean(f["encrypted"].ToString()) == true).ToList();

                            switch (table["table_name"].ToString().ToLower())
                            {
                                case "bnplan":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Bnplan",
                                                   new SqlParameter("@loads_id", loadsId),
                                                   new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                    new SqlParameter("@BenCode", node.Element("BenCode").Value),
                                                    new SqlParameter("@BenPlan", node.Element("BenPlan").Value),
                                                    new SqlParameter("@Active", node.Element("Active").Value),
                                                    new SqlParameter("@Mandatory", node.Element("Mandatory").Value),
                                                    new SqlParameter("@AutoSetup", node.Element("AutoSetup").Value),
                                                    new SqlParameter("@AutoEnroll", node.Element("AutoEnroll").Value),
                                                    new SqlParameter("@Beneficiaries", node.Element("Beneficiaries").Value),
                                                    new SqlParameter("@Dependents", node.Element("Dependents").Value),
                                                    new SqlParameter("@Funds", node.Element("Funds").Value),
                                                    new SqlParameter("@Vendor", node.Element("Vendor").Value),
                                                    new SqlParameter("@Description", node.Element("Description").Value),
                                                    new SqlParameter("@ShortDesc", node.Element("ShortDesc").Value),
                                                    new SqlParameter("@DateAdded", node.Element("DateAdded").Value),
                                                    new SqlParameter("@AddedBy", node.Element("AddedBy").Value),
                                                    new SqlParameter("@DateChanged", node.Element("DateChanged").Value),
                                                    new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value),
                                                    new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value),
                                                    new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value),
                                                    new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value),
                                                    new SqlParameter("@RefGroup", node.Element("RefGroup").Value),
                                                    new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value),
                                                    new SqlParameter("@MiscInt2", node.Element("MiscInt2").Value),
                                                    new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                                    new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                                    new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                                    new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                                    new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                                    new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                                    new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                                    new SqlParameter("@MiscAlpha3", node.Element("MiscAlpha3").Value),
                                                    new SqlParameter("@MiscAlpha4", node.Element("MiscAlpha4").Value),
                                                    new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                                    new SqlParameter("@UserAlpha2", node.Element("UserAlpha2").Value),
                                                    new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                                    new SqlParameter("@UserDate1", node.Element("UserDate1").Value),
                                                    new SqlParameter("@AllowSelfService", node.Element("AllowSelfService").Value));
                                    break;
                                case "bneplany":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Bneplany",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value),
                                                new SqlParameter("@PlanYr", node.Element("PlanYr").Value),
                                                new SqlParameter("@BenCode", node.Element("BenCode").Value),
                                                new SqlParameter("@BenPlan", node.Element("BenPlan").Value),
                                                new SqlParameter("@AmtsOk", node.Element("AmtsOk").Value),
                                                new SqlParameter("@AmtComment", node.Element("AmtComment").Value),
                                                new SqlParameter("@PlanYrStatus", node.Element("PlanYrStatus").Value),
                                                new SqlParameter("@StatusDate", node.Element("StatusDate").Value),
                                                new SqlParameter("@SelectDate", node.Element("SelectDate").Value),
                                                new SqlParameter("@ActiveDate", node.Element("ActiveDate").Value),
                                                new SqlParameter("@InactiveDate", node.Element("InactiveDate").Value),
                                                new SqlParameter("@Age", node.Element("Age").Value),
                                                new SqlParameter("@Exceptions", node.Element("Exceptions").Value),
                                                new SqlParameter("@ExCodeList", node.Element("ExCodeList").Value),
                                                new SqlParameter("@ExCodeDesc", node.Element("ExCodeDesc").Value),
                                                new SqlParameter("@PpPlanAmt", node.Element("PpPlanAmt").Value),
                                                new SqlParameter("@PpErPaidAmt", node.Element("PpErPaidAmt").Value),
                                                new SqlParameter("@PpEePaidAmt", node.Element("PpEePaidAmt").Value),
                                                new SqlParameter("@PpAddBackAmt", node.Element("PpAddBackAmt").Value),
                                                new SqlParameter("@PpEarnAmt", node.Element("PpEarnAmt").Value),
                                                new SqlParameter("@PpPercent", node.Element("PpPercent").Value),
                                                new SqlParameter("@PlanYrAmt", node.Element("PlanYrAmt").Value),
                                                new SqlParameter("@PlanErPaidAmt", node.Element("PlanErPaidAmt").Value),
                                                new SqlParameter("@PlanEePaidAmt", node.Element("PlanEePaidAmt").Value),
                                                new SqlParameter("@PlanAddBackAmt", node.Element("PlanAddBackAmt").Value),
                                                new SqlParameter("@PlanEarnAmt", node.Element("PlanEarnAmt").Value),
                                                new SqlParameter("@xLifeOthX", node.Element("xLifeOthX").Value),
                                                new SqlParameter("@CoverageAmt", node.Element("CoverageAmt").Value),
                                                new SqlParameter("@OtherValue", node.Element("OtherValue").Value),
                                                new SqlParameter("@xMedDentX", node.Element("xMedDentX").Value),
                                                new SqlParameter("@xDefContrbtnX", node.Element("xDefContrbtnX").Value),
                                                new SqlParameter("@EePercent", node.Element("EePercent").Value),
                                                new SqlParameter("@EePlanAmt", node.Element("EePlanAmt").Value),
                                                new SqlParameter("@EePpAmt", node.Element("EePpAmt").Value),
                                                new SqlParameter("@xBuySellVacX", node.Element("xBuySellVacX").Value),
                                                new SqlParameter("@HrsBought", node.Element("HrsBought").Value),
                                                new SqlParameter("@HrsSold", node.Element("HrsSold").Value),
                                                new SqlParameter("@CostPerHr", node.Element("CostPerHr").Value),
                                                new SqlParameter("@ExtendedCost", node.Element("ExtendedCost").Value),
                                                new SqlParameter("@AccrualCode", node.Element("AccrualCode").Value),
                                                new SqlParameter("@PaymentMethod", node.Element("PaymentMethod").Value),
                                                new SqlParameter("@PaymentMethodDesc", node.Element("PaymentMethodDesc").Value),
                                                new SqlParameter("@ManAccrGend", node.Element("ManAccrGend").Value),
                                                new SqlParameter("@xSavingsBondsX", node.Element("xSavingsBondsX").Value),
                                                new SqlParameter("@xReimbursementX", node.Element("xReimbursementX").Value),
                                                new SqlParameter("@DateAdded", node.Element("DateAdded").Value),
                                                new SqlParameter("@AddedBy", node.Element("AddedBy").Value),
                                                new SqlParameter("@DateChanged", node.Element("DateChanged").Value),
                                                new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value),
                                                new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value),
                                                new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value),
                                                new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value),
                                                new SqlParameter("@RefGroup", node.Element("RefGroup").Value),
                                                new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value),
                                                new SqlParameter("@MiscInt2", node.Element("MiscInt2").Value),
                                                new SqlParameter("@MiscInt3", node.Element("MiscInt3").Value),
                                                new SqlParameter("@MiscInt4", node.Element("MiscInt4").Value),
                                                new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                                new SqlParameter("@UserInt2", node.Element("UserInt2").Value),
                                                new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                                new SqlParameter("@MiscDec2", node.Element("MiscDec2").Value),
                                                new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                                new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                                new SqlParameter("@MiscLog2", node.Element("MiscLog2").Value),
                                                new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                                new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                                new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                                new SqlParameter("@MiscAlpha3", node.Element("MiscAlpha3").Value),
                                                new SqlParameter("@MiscAlpha4", node.Element("MiscAlpha4").Value),
                                                new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                                new SqlParameter("@UserAlpha2", node.Element("UserAlpha2").Value),
                                                new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                                new SqlParameter("@MiscDate2", node.Element("MiscDate2").Value),
                                                new SqlParameter("@UserDate1", node.Element("UserDate1").Value),
                                                new SqlParameter("@ContributionBeginsPayPeriod", node.Element("ContributionBeginsPayPeriod").Value));
                                    break;
                                case "bntype":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Bntype",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                new SqlParameter("@BenType", node.Element("BenType").Value),
                                                new SqlParameter("@BenClass", node.Element("BenClass").Value),
                                                new SqlParameter("@TypeOfDefContr", node.Element("TypeOfDefContr").Value),
                                                new SqlParameter("@Description", node.Element("Description").Value),
                                                new SqlParameter("@ShortDesc", node.Element("ShortDesc").Value),
                                                new SqlParameter("@RequPreTax", node.Element("RequPreTax").Value),
                                                new SqlParameter("@RequPreTaxDesc", node.Element("RequPreTaxDesc").Value),
                                                new SqlParameter("@RequPostTax", node.Element("RequPostTax").Value),
                                                new SqlParameter("@RequEr", node.Element("RequEr").Value),
                                                new SqlParameter("@RequErDesc", node.Element("RequErDesc").Value),
                                                new SqlParameter("@RequAddGross", node.Element("RequAddGross").Value),
                                                new SqlParameter("@RequAddGrossDesc", node.Element("RequAddGrossDesc").Value),
                                                new SqlParameter("@RequAddNet", node.Element("RequAddNet").Value),
                                                new SqlParameter("@RequRecEarnId", node.Element("RequRecEarnId").Value),
                                                new SqlParameter("@RequRecEarnIdDesc", node.Element("RequRecEarnIdDesc").Value),
                                                new SqlParameter("@BenEntryAppl", node.Element("BenEntryAppl").Value),
                                                new SqlParameter("@BenEntryType", node.Element("BenEntryType").Value),
                                                new SqlParameter("@BenEntryName", node.Element("BenEntryName").Value),
                                                new SqlParameter("@BenQueryAppl", node.Element("BenQueryAppl").Value),
                                                new SqlParameter("@BenQueryType", node.Element("BenQueryType").Value),
                                                new SqlParameter("@BenQueryName", node.Element("BenQueryName").Value),
                                                new SqlParameter("@BplanEntryMenuAppl", node.Element("BplanEntryMenuAppl").Value),
                                                new SqlParameter("@BplanEntryMenuType", node.Element("BplanEntryMenuType").Value),
                                                new SqlParameter("@BplanEntryMenuName", node.Element("BplanEntryMenuName").Value),
                                                new SqlParameter("@EbenEntryAppl", node.Element("EbenEntryAppl").Value),
                                                new SqlParameter("@EbenEntryType", node.Element("EbenEntryType").Value),
                                                new SqlParameter("@EbenEntryName", node.Element("EbenEntryName").Value),
                                                new SqlParameter("@EbenQueryAppl", node.Element("EbenQueryAppl").Value),
                                                new SqlParameter("@EbenQueryType", node.Element("EbenQueryType").Value),
                                                new SqlParameter("@EbenQueryName", node.Element("EbenQueryName").Value),
                                                new SqlParameter("@BplanQueryMenuAppl", node.Element("BplanQueryMenuAppl").Value),
                                                new SqlParameter("@BplanQueryMenuType", node.Element("BplanQueryMenuType").Value),
                                                new SqlParameter("@BplanQueryMenuName", node.Element("BplanQueryMenuName").Value),
                                                new SqlParameter("@DateAdded", node.Element("DateAdded").Value),
                                                new SqlParameter("@AddedBy", node.Element("AddedBy").Value),
                                                new SqlParameter("@DateChanged", node.Element("DateChanged").Value),
                                                new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value),
                                                new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value),
                                                new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value),
                                                new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value),
                                                new SqlParameter("@RefGroup", node.Element("RefGroup").Value),
                                                new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value),
                                                new SqlParameter("@MiscInt2", node.Element("MiscInt2").Value),
                                                new SqlParameter("@MiscInt3", node.Element("MiscInt3").Value),
                                                new SqlParameter("@MiscInt4", node.Element("MiscInt4").Value),
                                                new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                                new SqlParameter("@UserInt2", node.Element("UserInt2").Value),
                                                new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                                new SqlParameter("@MiscDec2", node.Element("MiscDec2").Value),
                                                new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                                new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                                new SqlParameter("@MiscLog2", node.Element("MiscLog2").Value),
                                                new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                                new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                                new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                                new SqlParameter("@MiscAlpha3", node.Element("MiscAlpha3").Value),
                                                new SqlParameter("@MiscAlpha4", node.Element("MiscAlpha4").Value),
                                                new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                                new SqlParameter("@UserAlpha2", node.Element("UserAlpha2").Value),
                                                new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                                new SqlParameter("@MiscDate2", node.Element("MiscDate2").Value),
                                                new SqlParameter("@UserDate1", node.Element("UserDate1").Value));
                                    break;
                                case "busunit":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_BusUnit",
                                                    new SqlParameter("@loads_id", loadsId),
                                                    new SqlParameter("@BusUnit", node.Element("BusUnit").Value),
                                                    new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                    new SqlParameter("@Description", node.Element("Description").Value),
                                                    new SqlParameter("@ShortDesc", node.Element("ShortDesc").Value),
                                                    new SqlParameter("@SvcChgAcct", node.Element("SvcChgAcct").Value),
                                                    new SqlParameter("@TrDiscAcct", node.Element("TrDiscAcct").Value),
                                                    new SqlParameter("@ArAcct", node.Element("ArAcct").Value),
                                                    new SqlParameter("@DisGivAcct", node.Element("DisGivAcct").Value),
                                                    new SqlParameter("@FrtOutAcct", node.Element("FrtOutAcct").Value),
                                                    new SqlParameter("@WrOffAcct", node.Element("WrOffAcct").Value),
                                                    new SqlParameter("@DepAcct", node.Element("DepAcct").Value),
                                                    new SqlParameter("@ApDiscAcct", node.Element("ApDiscAcct").Value),
                                                    new SqlParameter("@ApPreAcct", node.Element("ApPreAcct").Value),
                                                    new SqlParameter("@ApAcct", node.Element("ApAcct").Value),
                                                    new SqlParameter("@FacilityNo", node.Element("FacilityNo").Value),
                                                    new SqlParameter("@VatRegistrNo", node.Element("VatRegistrNo").Value),
                                                    new SqlParameter("@UsOrFc", node.Element("").Value),
                                                    new SqlParameter("@BillWashAcct", node.Element("").Value),
                                                    new SqlParameter("@Address1", node.Element("").Value),
                                                    new SqlParameter("@Address2", node.Element("").Value),
                                                    new SqlParameter("@City", node.Element("").Value),
                                                    new SqlParameter("@State", node.Element("").Value),
                                                    new SqlParameter("@ZipCode", node.Element("").Value),
                                                    new SqlParameter("@Country", node.Element("").Value),
                                                    new SqlParameter("@Entity", node.Element("Entity").Value),
                                                    new SqlParameter("@BankId", node.Element("BankId").Value),
                                                    new SqlParameter("@BankAcctNo", node.Element("BankAcctNo").Value),
                                                    new SqlParameter("@PyTaxArea", node.Element("PyTaxArea").Value),
                                                    new SqlParameter("@PyPayrollCode", node.Element("PyPayrollCode").Value),
                                                    new SqlParameter("@PyJobCode", node.Element("PyJobCode").Value),
                                                    new SqlParameter("@PyUnionCode", node.Element("PyUnionCode").Value),
                                                    new SqlParameter("@PyEmpStatus", node.Element("PyEmpStatus").Value),
                                                    new SqlParameter("@PyRecGroup", node.Element("PyRecGroup").Value),
                                                    new SqlParameter("@PySetupGroupId", node.Element("PySetupGroupId").Value),
                                                    new SqlParameter("@PyTaxLocation", node.Element("PyTaxLocation").Value),
                                                    new SqlParameter("@DateAdded", node.Element("DateAdded").Value),
                                                    new SqlParameter("@AddedBy", node.Element("AddedBy").Value),
                                                    new SqlParameter("@DateChanged", node.Element("DateChanged").Value),
                                                    new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value),
                                                    new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value),
                                                    new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value),
                                                    new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value),
                                                    new SqlParameter("@RefGroup", node.Element("RefGroup").Value),
                                                    new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value),
                                                    new SqlParameter("@MiscInt2", node.Element("MiscInt2").Value),
                                                    new SqlParameter("@MiscInt3", node.Element("MiscInt3").Value),
                                                    new SqlParameter("@MiscInt4", node.Element("MiscInt4").Value),
                                                    new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                                    new SqlParameter("@UserInt2", node.Element("UserInt2").Value),
                                                    new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                                    new SqlParameter("@MiscDec2", node.Element("MiscDec2").Value),
                                                    new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                                    new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                                    new SqlParameter("@MiscLog2", node.Element("MiscLog2").Value),
                                                    new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                                    new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                                    new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                                    new SqlParameter("@MiscAlpha3", node.Element("MiscAlpha3").Value),
                                                    new SqlParameter("@MiscAlpha4", node.Element("MiscAlpha4").Value),
                                                    new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                                    new SqlParameter("@UserAlpha2", node.Element("UserAlpha2").Value),
                                                    new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                                    new SqlParameter("@MiscDate2", node.Element("MiscDate2").Value),
                                                    new SqlParameter("@UserDate1", node.Element("UserDate1").Value),
                                                    new SqlParameter("@BusunitToChangeTo", node.Element("BusunitToChangeTo").Value),
                                                    new SqlParameter("@ListOfXferBusunits", node.Element("ListOfXferBusunits").Value),
                                                    new SqlParameter("@RQApprReqSameAsBusUnit", node.Element("RQApprReqSameAsBusUnit").Value),
                                                    new SqlParameter("@APApprReqSameAsBusUnit", node.Element("APApprReqSameAsBusUnit").Value));
                                    break;
                                case "empauto":

                                    break;
                                case "empded2":

                                    break;
                                case "empearna":

                                    break;
                                case "empgrp1":

                                    break;
                                case "empgrp2":

                                    break;
                                case "employee":

                                    break;
                                case "mainacct":

                                    break;
                                case "pperiod":

                                    break;
                                case "pyjobcd":

                                    break;
                                case "pyunion":

                                    break;
                                case "tcard2":

                                    break;
                                case "trx1099":

                                    break;
                                case "vendor":

                                    break;

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
