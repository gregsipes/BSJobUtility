﻿using BSJobBase;
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
            JobDescription = "Parses XML files into database tables";
            AppConfigSectionName = "SBSReportsLoad";

        }

        public override void ExecuteJob()
        {
            try
            {

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

                        List<XElement> nodes = new List<XElement>();

                        //this case statement replaces the replaces the where_conditions table. We ran into issues converting the sql strings into the Linq To XML queries,
                        //so for the sake of time, we moved the where clauses here
                        switch (table["table_name"].ToString().ToLower())
                        {
                            case "empded2":
                                nodes = xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").Where(n => n.Element("DeductCode").ToString().Contains("uf")).ToList();
                                break;
                            case "employee":
                                nodes = xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").Where(n => n.Element("CompanyId") != null && n.Element("CompanyId").ToString().Contains("BNEWS")).ToList();
                                break;
                            case "tcard2":
                                nodes = xml.Root.Elements("tt" + table["table_name"].ToString() + "Row").Where(n => Convert.ToDateTime(n.Element("TrxDate").Value) >= DateTime.Now.AddYears(-3)).ToList();
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
                                                new SqlParameter("@xLifeOthX", node.Element("XLifeOthX").Value),
                                                new SqlParameter("@CoverageAmt", node.Element("CoverageAmt").Value),
                                                new SqlParameter("@OtherValue", node.Element("OtherValue").Value),
                                                new SqlParameter("@xMedDentX", node.Element("XMedDentX").Value),
                                                new SqlParameter("@xDefContrbtnX", node.Element("XDefContrbtnX").Value),
                                                new SqlParameter("@EePercent", node.Element("EePercent").Value),
                                                new SqlParameter("@EePlanAmt", node.Element("EePlanAmt").Value),
                                                new SqlParameter("@EePpAmt", node.Element("EePpAmt").Value),
                                                new SqlParameter("@xBuySellVacX", node.Element("XBuySellVacX").Value),
                                                new SqlParameter("@HrsBought", node.Element("HrsBought").Value),
                                                new SqlParameter("@HrsSold", node.Element("HrsSold").Value),
                                                new SqlParameter("@CostPerHr", node.Element("CostPerHr").Value),
                                                new SqlParameter("@ExtendedCost", node.Element("ExtendedCost").Value),
                                                new SqlParameter("@AccrualCode", node.Element("AccrualCode").Value),
                                                new SqlParameter("@PaymentMethod", node.Element("PaymentMethod").Value),
                                                new SqlParameter("@PaymentMethodDesc", node.Element("PaymentMethodDesc").Value),
                                                new SqlParameter("@ManAccrGend", node.Element("ManAccrGend").Value),
                                                new SqlParameter("@xSavingsBondsX", node.Element("XSavingsBondsX").Value),
                                                new SqlParameter("@xReimbursementX", node.Element("XReimbursementX").Value),
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
                                                    new SqlParameter("@UsOrFc", node.Element("UsOrFc").Value),
                                                    new SqlParameter("@BillWashAcct", node.Element("BillWashAcct").Value),
                                                    new SqlParameter("@Address1", node.Element("Address1").Value),
                                                    new SqlParameter("@Address2", node.Element("Address2").Value),
                                                    new SqlParameter("@City", node.Element("City").Value),
                                                    new SqlParameter("@State", node.Element("State").Value),
                                                    new SqlParameter("@ZipCode", node.Element("ZipCode").Value),
                                                    new SqlParameter("@Country", node.Element("Country").Value),
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
                                                    new SqlParameter("@RQApprReqSameAsBusUnit", node.Element("RqApprReqSameAsBusUnit").Value),
                                                    new SqlParameter("@APApprReqSameAsBusUnit", node.Element("ApApprReqSameAsBusUnit").Value));
                                    break;
                                case "empauto":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Empauto",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value),
                                                new SqlParameter("@SeqNo", node.Element("SeqNo").Value),
                                                new SqlParameter("@EffectDate", node.Element("EffectDate").Value),
                                                new SqlParameter("@ExpireDate", node.Element("ExpireDate").Value),
                                                new SqlParameter("@Description", node.Element("Description").Value),
                                                new SqlParameter("@ShortDesc", node.Element("ShortDesc").Value),
                                                new SqlParameter("@PrenoteStatus", node.Element("PrenoteStatus").Value),
                                                new SqlParameter("@PrenoteStatusDesc", node.Element("PrenoteStatusDesc").Value),
                                                new SqlParameter("@PrenoteDate", node.Element("PrenoteDate").Value),
                                                new SqlParameter("@AchBankTransit", node.Element("AchBankTransit").Value),
                                                new SqlParameter("@AchBankAcct", node.Element("AchBankAcct").Value),
                                                new SqlParameter("@AchAcctType", node.Element("AchAcctType").Value),
                                                new SqlParameter("@AchAcctTypeDesc", node.Element("AchAcctTypeDesc").Value),
                                                new SqlParameter("@FixedAmt", node.Element("FixedAmt").Value),
                                                new SqlParameter("@Percent", node.Element("Percent").Value),
                                                new SqlParameter("@FreqCode", node.Element("FreqCode").Value),
                                                new SqlParameter("@MaxNoOfPay", node.Element("MaxNoOfPay").Value),
                                                new SqlParameter("@NoOfPayProc", node.Element("NoOfPayProc").Value),
                                                new SqlParameter("@MaxAmt", node.Element("MaxAmt").Value),
                                                new SqlParameter("@AmtProcessed", node.Element("AmtProcessed").Value),
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
                                                new SqlParameter("@ExcludedCheckTypes", node.Element("ExcludedCheckTypes").Value));
                                    break;
                                case "empded2":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Empded2",
                                               new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value),
                                                new SqlParameter("@DeductCode", node.Element("DeductCode").Value),
                                                new SqlParameter("@EffectDate", node.Element("EffectDate").Value),
                                                new SqlParameter("@ExpireDate", node.Element("ExpireDate").Value),
                                                new SqlParameter("@CalcSeq", node.Element("CalcSeq").Value),
                                                new SqlParameter("@Active", node.Element("Active").Value),
                                                new SqlParameter("@BenCode", node.Element("BenCode").Value),
                                                new SqlParameter("@BenPlan", node.Element("BenPlan").Value),
                                                new SqlParameter("@PlanYr", node.Element("PlanYr").Value),
                                                new SqlParameter("@CalcType", node.Element("CalcType").Value),
                                                new SqlParameter("@CalcTypeDesc", node.Element("CalcTypeDesc").Value),
                                                new SqlParameter("@FixedAmt", node.Element("FixedAmt").Value),
                                                new SqlParameter("@Percent", node.Element("Percent").Value),
                                                new SqlParameter("@AddlFixedAmt", node.Element("AddlFixedAmt").Value),
                                                new SqlParameter("@TableId", node.Element("TableId").Value),
                                                new SqlParameter("@LimitType", node.Element("LimitType").Value),
                                                new SqlParameter("@LimitTypeDesc", node.Element("LimitTypeDesc").Value),
                                                new SqlParameter("@LimitAmt", node.Element("LimitAmt").Value),
                                                new SqlParameter("@FreqCode", node.Element("FreqCode").Value),
                                                new SqlParameter("@BasisId", node.Element("BasisId").Value),
                                                new SqlParameter("@PreTaxBasisId", node.Element("PreTaxBasisId").Value),
                                                new SqlParameter("@GarnishId", node.Element("GarnishId").Value),
                                                new SqlParameter("@AllocId", node.Element("AllocId").Value),
                                                new SqlParameter("@EligStatus", node.Element("EligStatus").Value),
                                                new SqlParameter("@EligHrBasis", node.Element("EligHrBasis").Value),
                                                new SqlParameter("@EligHrs", node.Element("EligHrs").Value),
                                                new SqlParameter("@BonusPct", node.Element("BonusPct").Value),
                                                new SqlParameter("@CommPct", node.Element("CommPct").Value),
                                                new SqlParameter("@EarningsGroup", node.Element("EarningsGroup").Value),
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
                                                new SqlParameter("@ExternalReference", node.Element("ExternalReference").Value),
                                                new SqlParameter("@TempDeductCode", node.Element("TempDeductCode").Value),
                                                new SqlParameter("@EditInBenefitsOnly", node.Element("EditInBenefitsOnly").Value),
                                                new SqlParameter("@BonusAndCommTableIDs", node.Element("BonusAndCommTableIDs").Value),
                                                new SqlParameter("@GarnishVendor", node.Element("GarnishVendor").Value),
                                                new SqlParameter("@GarnishInvDesc", node.Element("GarnishInvDesc").Value),
                                                new SqlParameter("@GarnishText", node.Element("GarnishText").Value),
                                                new SqlParameter("@LimitBasisID", node.Element("LimitBasisID").Value));
                                    break;
                                case "empearna":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Empearna",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value),
                                                new SqlParameter("@CalYear", node.Element("CalYear").Value),
                                                new SqlParameter("@EarningsCode", node.Element("EarningsCode").Value),
                                                new SqlParameter("@TaxArea", node.Element("TaxArea").Value),
                                                new SqlParameter("@Hrs", node.Element("Hrs").Value),
                                                new SqlParameter("@TotalHrsForYr", node.Element("TotalHrsForYr").Value),
                                                new SqlParameter("@Amts", node.Element("Amts").Value),
                                                new SqlParameter("@TotalAmtForYr", node.Element("TotalAmtForYr").Value),
                                                new SqlParameter("@DiffAmts", node.Element("DiffAmts").Value),
                                                new SqlParameter("@TotalDiffAmtForYr", node.Element("TotalDiffAmtForYr").Value),
                                                new SqlParameter("@DateAdded", node.Element("DateAdded").Value),
                                                new SqlParameter("@AddedBy", node.Element("AddedBy").Value),
                                                new SqlParameter("@DateChanged", node.Element("DateChanged").Value),
                                                new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value),
                                                new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value),
                                                new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value),
                                                new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value),
                                                new SqlParameter("@RefGroup", node.Element("RefGroup").Value),
                                                new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value),
                                                new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                                new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                                new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                                new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                                new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                                new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                                new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                                new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                                new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                                new SqlParameter("@UserDate1", node.Element("UserDate1").Value));
                                    break;
                                case "empgrp1":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Empgrp1",
                                                new SqlParameter("@loads_id", loadsId),
                                                new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                                new SqlParameter("@GroupId", node.Element("GroupId").Value),
                                                new SqlParameter("@Description", node.Element("Description").Value),
                                                new SqlParameter("@ShortDesc", node.Element("ShortDesc").Value),
                                                new SqlParameter("@AllowManual", node.Element("AllowManual").Value),
                                                new SqlParameter("@Comments", node.Element("Comments").Value),
                                                new SqlParameter("@DateAdded", node.Element("DateAdded").Value),
                                                new SqlParameter("@AddedBy", node.Element("AddedBy").Value),
                                                new SqlParameter("@DateChanged", node.Element("DateChanged").Value),
                                                new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value),
                                                new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value),
                                                new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value),
                                                new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value),
                                                new SqlParameter("@RefGroup", node.Element("RefGroup").Value),
                                                new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value),
                                                new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                                new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                                new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                                new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                                new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                                new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                                new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                                new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                                new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                                new SqlParameter("@UserDate1", node.Element("UserDate1").Value),
                                                new SqlParameter("@SortField1", node.Element("SortField1").Value),
                                                new SqlParameter("@SortField2", node.Element("SortField2").Value),
                                                new SqlParameter("@SortField3", node.Element("SortField3").Value),
                                                new SqlParameter("@SortField4", node.Element("SortField4").Value),
                                                new SqlParameter("@SortField5", node.Element("SortField5").Value),
                                                new SqlParameter("@GroupUsedForTA", node.Element("GroupUSedForTA").Value));
                                    break;
                                case "empgrp2":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Empgrp2",
                                            new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                            new SqlParameter("@GroupId", node.Element("GroupId").Value),
                                            new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value),
                                            new SqlParameter("@Alpha", node.Element("Alpha").Value),
                                            new SqlParameter("@BusUnit", node.Element("BusUnit").Value),
                                            new SqlParameter("@MemberSource", node.Element("MemberSource").Value),
                                            new SqlParameter("@MemberSourceDesc", node.Element("MemberSourceDesc").Value),
                                            new SqlParameter("@AllowTrx", node.Element("AllowTrx").Value),
                                            new SqlParameter("@DateAdded", node.Element("DateAdded").Value),
                                            new SqlParameter("@AddedBy", node.Element("AddedBy").Value),
                                            new SqlParameter("@DateChanged", node.Element("DateChanged").Value),
                                            new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value),
                                            new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value),
                                            new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value),
                                            new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value),
                                            new SqlParameter("@RefGroup", node.Element("RefGroup").Value),
                                            new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value),
                                            new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                            new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                            new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                            new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                            new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                            new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                            new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                            new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                            new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                            new SqlParameter("@UserDate1", node.Element("UserDate1").Value));
                                    break;
                                case "employee":


                                           var x = new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value);
                                            x = new SqlParameter("@EmployeeId", node.Element("EmployeeId").Value);
                                            x = new SqlParameter("@PersonalID", node.Element("PersonalId").Value);
                                            x = new SqlParameter("@Alpha", node.Element("Alpha").Value);
                                            x = new SqlParameter("@FirstName", node.Element("FirstName").Value);
                                            x = new SqlParameter("@MiddleInit", node.Element("MiddleInit").Value);
                                            x = new SqlParameter("@LastName", node.Element("LastName").Value);
                                            x = new SqlParameter("@UserId", node.Element("PseudoUserID").Value);
                                            x = new SqlParameter("@CustomerNo", node.Element("CustomerNo").Value);
                                            x = new SqlParameter("@Vendor", node.Element("Vendor").Value);
                                            x = new SqlParameter("@FacilityNo", node.Element("FacilityNo").Value);
                                            x = new SqlParameter("@JobClass", node.Element("JobClass").Value);
                                            x = new SqlParameter("@JobCode", node.Element("JobCode").Value);
                                            x = new SqlParameter("@Position", node.Element("Position").Value);
                                            x = new SqlParameter("@BusUnit", node.Element("BusUnit").Value);
                                            x = new SqlParameter("@CompanyId", node.Element("CompanyId").Value);
                                            x = new SqlParameter("@Project", node.Element("Project").Value);
                                            x = new SqlParameter("@ProjectWb", node.Element("ProjectWb").Value);
                                            x = new SqlParameter("@EarningsCode", node.Element("EarningsCode").Value);
                                            x = new SqlParameter("@ActualEmployee", node.Element("ActualEmployee").Value);
                                            x = new SqlParameter("@EmpStatus", node.Element("EmpStatus").Value);
                                            x = new SqlParameter("@HighCompEmp", node.Element("HighCompEmp").Value);
                                            x = new SqlParameter("@KeyEmployee", node.Element("KeyEmployee").Value);
                                            x = new SqlParameter("@Officer", node.Element("Officer").Value);
                                            x = new SqlParameter("@Tipped", node.Element("Tipped").Value);
                                            x = new SqlParameter("@PrimaryStateTableId", node.Element("PrimaryStateTableId").Value);
                                            x = new SqlParameter("@PrimaryStateFilingStatus", node.Element("PrimaryStateFilingStatus").Value);
                                            x = new SqlParameter("@PrimaryStateExemptions", node.Element("PrimaryStateExemptions").Value);
                                            x = new SqlParameter("@PrimaryLocalTableId", node.Element("PrimaryLocalTableId").Value);
                                            x = new SqlParameter("@PrimaryLocalFilingStatus", node.Element("PrimaryLocalFilingStatus").Value);
                                            x = new SqlParameter("@PrimaryLocalExemptions", node.Element("PrimaryLocalExemptions").Value);
                                            x = new SqlParameter("@FedTableId", node.Element("FedTableId").Value);
                                            x = new SqlParameter("@FedFilingStatus", node.Element("FedFilingStatus").Value);
                                            x = new SqlParameter("@FedExemptions", node.Element("FedExemptions").Value);
                                            x = new SqlParameter("@Name", node.Element("Name").Value);
                                            x = new SqlParameter("@Address1", node.Element("Address1").Value);
                                            x = new SqlParameter("@Address2", node.Element("Address2").Value);
                                            x = new SqlParameter("@City", node.Element("City").Value);
                                            x = new SqlParameter("@State", node.Element("State").Value);
                                            x = new SqlParameter("@ZipCode", node.Element("ZipCode").Value);
                                            x = new SqlParameter("@Country", node.Element("Country").Value);
                                            x = new SqlParameter("@Phone", node.Element("Phone").Value);
                                            x = new SqlParameter("@EmailAddress", node.Element("EmailAddress").Value);
                                            x = new SqlParameter("@Hourly", node.Element("Hourly").Value);
                                            x = new SqlParameter("@HourlyDesc", node.Element("HourlyDesc").Value);
                                            x = new SqlParameter("@Salary", node.Element("Salary").Value);
                                            x = new SqlParameter("@PayRate", node.Element("PayRate").Value);
                                            x = new SqlParameter("@BirthDate", node.Element("BirthDate").Value);
                                            x = new SqlParameter("@HireDate", node.Element("HireDate").Value);
                                            x = new SqlParameter("@ReHireDate", node.Element("ReHireDate").Value);
                                            x = new SqlParameter("@AnnivDate", node.Element("AnnivDate").Value);
                                            x = new SqlParameter("@TermDate", node.Element("TermDate").Value);
                                            x = new SqlParameter("@Terminated", node.Element("Terminated").Value);
                                            x = new SqlParameter("@ReasonCode", node.Element("ReasonCode").Value);
                                            x = new SqlParameter("@AdjHireDate", node.Element("AdjHireDate").Value);
                                            x = new SqlParameter("@SeniorityDate", node.Element("SeniorityDate").Value);
                                            x = new SqlParameter("@CanRoeDate", node.Element("CanRoeDate").Value);
                                            x = new SqlParameter("@TermComments", node.Element("TermComments").Value);
                                            x = new SqlParameter("@Sex", node.Element("Sex").Value);
                                            x = new SqlParameter("@SexDesc", node.Element("SexDesc").Value);
                                            x = new SqlParameter("@EeocClass", node.Element("EeocClass").Value);
                                            x = new SqlParameter("@EeocClassDesc", node.Element("EeocClassDesc").Value);
                                            x = new SqlParameter("@Handicapped", node.Element("Handicapped").Value);
                                            x = new SqlParameter("@OffPhone", node.Element("OffPhone").Value);
                                            x = new SqlParameter("@OffPhoneExt", node.Element("OffPhoneExt").Value);
                                            x = new SqlParameter("@MaritalStat", node.Element("MaritalStat").Value);
                                            x = new SqlParameter("@MaritalStatDesc", node.Element("MaritalStatDesc").Value);
                                            x = new SqlParameter("@Exempt", node.Element("Exempt").Value);
                                            x = new SqlParameter("@PayrollCode", node.Element("PayrollCode").Value);
                                            x = new SqlParameter("@UnionCode", node.Element("UnionCode").Value);
                                            x = new SqlParameter("@TaxArea", node.Element("TaxArea").Value);
                                            x = new SqlParameter("@ShiftNo", node.Element("ShiftNo").Value);
                                            x = new SqlParameter("@RecGroup", node.Element("RecGroup").Value);
                                            x = new SqlParameter("@ExpenseMain", node.Element("ExpenseMain").Value);
                                            x = new SqlParameter("@PcheckMsg", node.Element("PcheckMsg").Value);
                                            x = new SqlParameter("@PcheckMsgExp", node.Element("PcheckMsgExp").Value);
                                            x = new SqlParameter("@BenPackage", node.Element("BenPackage").Value);
                                            x = new SqlParameter("@Cobra", node.Element("Cobra").Value);
                                            x = new SqlParameter("@CobraComplete", node.Element("CobraComplete").Value);
                                            x = new SqlParameter("@CobraSrcEmp", node.Element("CobraSrcEmp").Value);
                                            x = new SqlParameter("@CobraSrcDepId", node.Element("CobraSrcDepId").Value);
                                            x = new SqlParameter("@CobraEvent", node.Element("CobraEvent").Value);
                                            x = new SqlParameter("@CobraEventDate", node.Element("CobraEventDate").Value);
                                            x = new SqlParameter("@CobraEeNoteDate", node.Element("CobraEeNoteDate").Value);
                                            x = new SqlParameter("@CobraRightsDate", node.Element("CobraRightsDate").Value);
                                            x = new SqlParameter("@CobraElectCode", node.Element("CobraElectCode").Value);
                                            x = new SqlParameter("@CobraElectCodeDesc", node.Element("CobraElectCodeDesc").Value);
                                            x = new SqlParameter("@CobraElectDate", node.Element("CobraElectDate").Value);
                                            x = new SqlParameter("@CobraLatePayNote", node.Element("CobraLatePayNote").Value);
                                            x = new SqlParameter("@CobraTermLetDate", node.Element("CobraTermLetDate").Value);
                                            x = new SqlParameter("@CobraTermCode", node.Element("CobraTermCode").Value);
                                            x = new SqlParameter("@CobraTermCodeDesc", node.Element("CobraTermCodeDesc").Value);
                                            x = new SqlParameter("@CobraTermDate", node.Element("CobraTermDate").Value);
                                            x = new SqlParameter("@CobraFirstPeriod", node.Element("CobraFirstPeriod").Value);
                                            x = new SqlParameter("@CobraLastPeriod", node.Element("CobraLastPeriod").Value);
                                            x = new SqlParameter("@CobraStartDate", node.Element("CobraStartDate").Value);
                                            x = new SqlParameter("@CobraEndDate", node.Element("CobraEndDate").Value);
                                            x = new SqlParameter("@W2Flag", node.Element("W2Flag").Value);
                                            x = new SqlParameter("@W2MiscValue", node.Element("W2MiscValue").Value);
                                            x = new SqlParameter("@Veteran", node.Element("Veteran").Value);
                                            x = new SqlParameter("@VeteranDesc", node.Element("VeteranDesc").Value);
                                            x = new SqlParameter("@Rank", node.Element("Rank").Value);
                                            x = new SqlParameter("@Branch", node.Element("Branch").Value);
                                            x = new SqlParameter("@DateDischarged", node.Element("DateDischarged").Value);
                                            x = new SqlParameter("@DischargeType", node.Element("DischargeType").Value);
                                            x = new SqlParameter("@DischargeTypeDesc", node.Element("DischargeTypeDesc").Value);
                                            x = new SqlParameter("@CurrentService", node.Element("CurrentService").Value);
                                            x = new SqlParameter("@CurrentServiceDesc", node.Element("CurrentServiceDesc").Value);
                                            x = new SqlParameter("@PassportNo", node.Element("PassportNo").Value);
                                            x = new SqlParameter("@PassportExpDate", node.Element("PassportExpDate").Value);
                                            x = new SqlParameter("@CountryCitizen", node.Element("CountryCitizen").Value);
                                            x = new SqlParameter("@BirthPlace", node.Element("BirthPlace").Value);
                                            x = new SqlParameter("@VerifyType", node.Element("VerifyType").Value);
                                            x = new SqlParameter("@VerifyDate", node.Element("VerifyDate").Value);
                                            x = new SqlParameter("@ExpireDate", node.Element("ExpireDate").Value);
                                            x = new SqlParameter("@AlienNo", node.Element("AlienNo").Value);
                                            x = new SqlParameter("@AdmissionNo", node.Element("AdmissionNo").Value);
                                            x = new SqlParameter("@ListAType", node.Element("ListAType").Value);
                                            x = new SqlParameter("@ListADocNo", node.Element("ListADocNo").Value);
                                            x = new SqlParameter("@ListAExpDate", node.Element("ListAExpDate").Value);
                                            x = new SqlParameter("@ListBType", node.Element("ListBType").Value);
                                            x = new SqlParameter("@ListBDocNo", node.Element("ListBDocNo").Value);
                                            x = new SqlParameter("@ListBExpDate", node.Element("ListBExpDate").Value);
                                            x = new SqlParameter("@ListBState", node.Element("ListBState").Value);
                                            x = new SqlParameter("@ListBOther", node.Element("ListBOther").Value);
                                            x = new SqlParameter("@ListCType", node.Element("ListCType").Value);
                                            x = new SqlParameter("@ListCDocNo", node.Element("ListCDocNo").Value);
                                            x = new SqlParameter("@ListCExpDate", node.Element("ListCExpDate").Value);
                                            x = new SqlParameter("@ListCInsForm", node.Element("ListCInsForm").Value);
                                            x = new SqlParameter("@ApprovalDate", node.Element("ApprovalDate").Value);
                                            x = new SqlParameter("@ApprovedBy", node.Element("ApprovedBy").Value);
                                            x = new SqlParameter("@ApprovedByTitle", node.Element("ApprovedByTitle").Value);
                                            x = new SqlParameter("@ReviewDate", node.Element("ReviewDate").Value);
                                            x = new SqlParameter("@EmergenyContact", node.Element("EmergenyContact").Value);
                                            x = new SqlParameter("@EmergAddress1", node.Element("EmergAddress1").Value);
                                            x = new SqlParameter("@EmergAddress2", node.Element("EmergAddress2").Value);
                                            x = new SqlParameter("@EmergCity", node.Element("EmergCity").Value);
                                            x = new SqlParameter("@EmergState", node.Element("EmergState").Value);
                                            x = new SqlParameter("@EmergZipCode", node.Element("EmergZipCode").Value);
                                            x = new SqlParameter("@EmergRelation", node.Element("EmergRelation").Value);
                                            x = new SqlParameter("@EmergPrimPhoneNo", node.Element("EmergPrimPhoneNo").Value);
                                            x = new SqlParameter("@EmergSecPhoneNo", node.Element("EmergSecPhoneNo").Value);
                                            x = new SqlParameter("@Physician", node.Element("Physician").Value);
                                            x = new SqlParameter("@PhysicianPhoneNo", node.Element("PhysicianPhoneNo").Value);
                                            x = new SqlParameter("@NextPhysical", node.Element("NextPhysical").Value);
                                            x = new SqlParameter("@LastPhysical", node.Element("LastPhysical").Value);
                                            x = new SqlParameter("@PhyRenewMonth", node.Element("PhyRenewMonth").Value);
                                            x = new SqlParameter("@DonationDate", node.Element("DonationDate").Value);
                                            x = new SqlParameter("@PhysicalResult", node.Element("PhysicalResult").Value);
                                            x = new SqlParameter("@WorkRest", node.Element("WorkRest").Value);
                                            x = new SqlParameter("@MedicalCode", node.Element("MedicalCode").Value);
                                            x = new SqlParameter("@Height", node.Element("Height").Value);
                                            x = new SqlParameter("@Weight", node.Element("Weight").Value);
                                            x = new SqlParameter("@BloodType", node.Element("BloodType").Value);
                                            x = new SqlParameter("@RhFactor", node.Element("RhFactor").Value);
                                            x = new SqlParameter("@NoFedEx", node.Element("NoFedEx").Value);
                                            x = new SqlParameter("@NoStateEx", node.Element("NoStateEx").Value);
                                            x = new SqlParameter("@NoLocalEx", "");
                                            x = new SqlParameter("@FedFilingStat", node.Element("FedFilingStat").Value);
                                            x = new SqlParameter("@StateFilingStat", node.Element("StateFilingStat").Value);
                                            x = new SqlParameter("@LocalFilingStat", node.Element("SecondLocalFilingStatus").Value);
                                            x = new SqlParameter("@AllowReimb", node.Element("AllowReimb").Value);
                                            x = new SqlParameter("@NoOfFtes", node.Element("NoOfFtes").Value);
                                            x = new SqlParameter("@FteClass", node.Element("FteClass").Value);
                                            x = new SqlParameter("@DateAdded", node.Element("DateAdded").Value);
                                            x = new SqlParameter("@AddedBy", node.Element("AddedBy").Value);
                                            x = new SqlParameter("@DateChanged", node.Element("DateChanged").Value);
                                            x = new SqlParameter("@ChangedBy", node.Element("ChangedBy").Value);
                                            x = new SqlParameter("@XmitStatus", node.Element("XmitStatus").Value);
                                            x = new SqlParameter("@OkToDelete", node.Element("OkToDelete").Value);
                                            x = new SqlParameter("@AccessGroups", node.Element("AccessGroups").Value);
                                            x = new SqlParameter("@RefGroup", node.Element("RefGroup").Value);
                                            x = new SqlParameter("@MiscInt1", node.Element("MiscInt1").Value);
                                            x = new SqlParameter("@MiscInt2", node.Element("MiscInt2").Value);
                                            x = new SqlParameter("@MiscInt3", node.Element("MiscInt3").Value);
                                            x = new SqlParameter("@MiscInt4", node.Element("MiscInt4").Value);
                                            x = new SqlParameter("@UserInt1", node.Element("UserInt1").Value);
                                            x = new SqlParameter("@UserInt2", node.Element("UserInt2").Value);
                                            x = new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value);
                                            x = new SqlParameter("@MiscDec2", node.Element("MiscDec2").Value);
                                            x = new SqlParameter("@UserDec1", node.Element("UserDec1").Value);
                                            x = new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value);
                                            x = new SqlParameter("@MiscLog2", node.Element("MiscLog2").Value);
                                            x = new SqlParameter("@UserLog1", node.Element("UserLog1").Value);
                                            x = new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value);
                                            x = new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value);
                                            x = new SqlParameter("@MiscAlpha3", node.Element("MiscAlpha3").Value);
                                            x = new SqlParameter("@MiscAlpha4", node.Element("MiscAlpha4").Value);
                                            x = new SqlParameter("@MiscAlpha5", node.Element("MiscAlpha5").Value);
                                            x = new SqlParameter("@MiscAlpha6", node.Element("MiscAlpha6").Value);
                                            x = new SqlParameter("@MiscAlpha7", node.Element("MiscAlpha7").Value);
                                            x = new SqlParameter("@MiscAlpha8", node.Element("MiscAlpha8").Value);
                                            x = new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value);
                                            x = new SqlParameter("@UserAlpha2", node.Element("UserAlpha2").Value);
                                            x = new SqlParameter("@UserAlpha3", node.Element("UserAlpha3").Value);
                                            x = new SqlParameter("@UserAlpha4", node.Element("UserAlpha4").Value);
                                            x = new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value);
                                            x = new SqlParameter("@MiscDate2", node.Element("MiscDate2").Value);
                                            x = new SqlParameter("@UserDate1", node.Element("UserDate1").Value);
                                            x = new SqlParameter("@ListATypeDesc", node.Element("ListATypeDesc").Value);
                                            x = new SqlParameter("@Supervisor", node.Element("SupervisorId").Value);
                                            x = new SqlParameter("@ListCTypeDesc", node.Element("ListCTypeDesc").Value);
                                            x = new SqlParameter("@UITaxArea", node.Element("UiTaxArea").Value);
                                            x = new SqlParameter("@ReHireDate", node.Element("ReHireDate").Value);
                                            x = new SqlParameter("@VerificationTypeDesc", node.Element("VerificationTypeDesc").Value);
                                            x = new SqlParameter("@EmployeeFTEs", node.Element("EmployeeFtes").Value);
                                            x = new SqlParameter("@AllowPaychecksAfterTerm", node.Element("AllowPaychecksAfterTerm").Value);
                                            x = new SqlParameter("@ListBTypeDesc", node.Element("ListBTypeDesc").Value);
                                            x = new SqlParameter("@ListAOther", node.Element("ListAOther").Value);
                                            x = new SqlParameter("@SortField1", node.Element("SortField1").Value);
                                            x = new SqlParameter("@SortField2", node.Element("SortField2").Value);
                                            x = new SqlParameter("@SortField3", node.Element("SortField3").Value);
                                            x = new SqlParameter("@SortField4", node.Element("SortField4").Value);
                                            x = new SqlParameter("@SortField5", node.Element("SortField5").Value);
                                            x = new SqlParameter("@AllowDepartmentTransfer", node.Element("AllowDepartmentTransfer").Value);
                                            x = new SqlParameter("@AllowJobTransfer", node.Element("AllowJobTransfer").Value);
                                            x = new SqlParameter("@AllowPunchAcrossDays", node.Element("AllowPunchAcrossDays").Value);
                                            x = new SqlParameter("@ListOfValidDepartments", node.Element("ListOfValidDepartments").Value);
                                            x = new SqlParameter("@ListofValidJobs", node.Element("ListofValidJobs").Value);
                                            x = new SqlParameter("@SupervisorId", node.Element("SupervisorId").Value);
                                            x = new SqlParameter("@WorkEmailAddress", node.Element("WorkEmailAddress").Value);
                                            x = new SqlParameter("@PseudoUserID", node.Element("PseudoUserID").Value);
                                            x = new SqlParameter("@TimeKeepingProfile", node.Element("TimeKeepingProfile").Value);
                                            x = new SqlParameter("@ScheduleID", node.Element("ScheduleID").Value);
                                            x = new SqlParameter("@SchedulingMethod", node.Element("SchedulingMethod").Value);
                                            x = new SqlParameter("@Team", node.Element("Team").Value);
                                            x = new SqlParameter("@SupervisoryEE", node.Element("SupervisoryEE").Value);
                                            x = new SqlParameter("@PsGroupNo", node.Element("PsGroupNo").Value);
                                            x = new SqlParameter("@LastEvalDateAndMiltime", node.Element("LastEvalDateAndMiltime").Value);
                                            x = new SqlParameter("@CellPhone", node.Element("CellPhone").Value);
                                            x = new SqlParameter("@NameSuffix", node.Element("NameSuffix").Value);
                                            x = new SqlParameter("@OverrideUIDeductCode", node.Element("OverrideUIDeductCode").Value);
                                            x = new SqlParameter("@NewHireBenefitsProcessed", node.Element("NewHireBenefitsProcessed").Value);
                                            x = new SqlParameter("@Eligible", node.Element("Eligible").Value);
                                            x = new SqlParameter("@GeographicCode", node.Element("GeographicCode").Value);
                                            x = new SqlParameter("@Rehire", node.Element("Rehire").Value);
                                            x = new SqlParameter("@MiscDec3", node.Element("MiscDec3").Value);
                                            x = new SqlParameter("@MiscDec4", node.Element("MiscDec4").Value);
                                            x = new SqlParameter("@MiscLog3", node.Element("MiscLog3").Value);
                                            x = new SqlParameter("@MiscLog4", node.Element("MiscLog4").Value);
                                            x = new SqlParameter("@MiscAlpha9", node.Element("MiscAlpha9").Value);
                                            x = new SqlParameter("@MiscAlpha10", node.Element("MiscAlpha10").Value);
                                            x = new SqlParameter("@MiscAlpha11", node.Element("MiscAlpha11").Value);
                                            x = new SqlParameter("@MiscAlpha12", node.Element("MiscAlpha12").Value);
                                            x = new SqlParameter("@MiscAlpha13", node.Element("MiscAlpha13").Value);
                                            x = new SqlParameter("@MiscAlpha14", node.Element("MiscAlpha14").Value);
                                            x = new SqlParameter("@MiscDate3", node.Element("MiscDate3").Value);
                                            x = new SqlParameter("@EESpecificTimeEntryTemplate", node.Element("EESpecificTimeEntryTemplate").Value);
                                    x = new SqlParameter("@TimeEntryTemplate", node.Element("TimeEntryTemplate").Value);

                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Employee",
                                            new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value),
                                            new SqlParameter("@EmployeeId", node.Element("EmployeeId").Value),
                                            new SqlParameter("@PersonalID", node.Element("PersonalId").Value),
                                            new SqlParameter("@Alpha", node.Element("Alpha").Value),
                                            new SqlParameter("@FirstName", node.Element("FirstName").Value),
                                            new SqlParameter("@MiddleInit", node.Element("MiddleInit").Value),
                                            new SqlParameter("@LastName", node.Element("LastName").Value),
                                            new SqlParameter("@UserId", node.Element("PseudoUserID").Value),
                                            new SqlParameter("@CustomerNo", node.Element("CustomerNo").Value),
                                            new SqlParameter("@Vendor", node.Element("Vendor").Value),
                                            new SqlParameter("@FacilityNo", node.Element("FacilityNo").Value),
                                            new SqlParameter("@JobClass", node.Element("JobClass").Value),
                                            new SqlParameter("@JobCode", node.Element("JobCode").Value),
                                            new SqlParameter("@Position", node.Element("Position").Value),
                                            new SqlParameter("@BusUnit", node.Element("BusUnit").Value),
                                            new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                            new SqlParameter("@Project", node.Element("Project").Value),
                                            new SqlParameter("@ProjectWb", node.Element("ProjectWb").Value),
                                            new SqlParameter("@EarningsCode", node.Element("EarningsCode").Value),
                                            new SqlParameter("@ActualEmployee", node.Element("ActualEmployee").Value),
                                            new SqlParameter("@EmpStatus", node.Element("EmpStatus").Value),
                                            new SqlParameter("@HighCompEmp", node.Element("HighCompEmp").Value),
                                            new SqlParameter("@KeyEmployee", node.Element("KeyEmployee").Value),
                                            new SqlParameter("@Officer", node.Element("Officer").Value),
                                            new SqlParameter("@Tipped", node.Element("Tipped").Value),
                                            new SqlParameter("@PrimaryStateTableId", node.Element("PrimaryStateTableId").Value),
                                            new SqlParameter("@PrimaryStateFilingStatus", node.Element("PrimaryStateFilingStatus").Value),
                                            new SqlParameter("@PrimaryStateExemptions", node.Element("PrimaryStateExemptions").Value),
                                            new SqlParameter("@PrimaryLocalTableId", node.Element("PrimaryLocalTableId").Value),
                                            new SqlParameter("@PrimaryLocalFilingStatus", node.Element("PrimaryLocalFilingStatus").Value),
                                            new SqlParameter("@PrimaryLocalExemptions", node.Element("PrimaryLocalExemptions").Value),
                                            new SqlParameter("@FedTableId", node.Element("FedTableId").Value),
                                            new SqlParameter("@FedFilingStatus", node.Element("FedFilingStatus").Value),
                                            new SqlParameter("@FedExemptions", node.Element("FedExemptions").Value),
                                            new SqlParameter("@Name", node.Element("Name").Value),
                                            new SqlParameter("@Address1", node.Element("Address1").Value),
                                            new SqlParameter("@Address2", node.Element("Address2").Value),
                                            new SqlParameter("@City", node.Element("City").Value),
                                            new SqlParameter("@State", node.Element("State").Value),
                                            new SqlParameter("@ZipCode", node.Element("ZipCode").Value),
                                            new SqlParameter("@Country", node.Element("Country").Value),
                                            new SqlParameter("@Phone", node.Element("Phone").Value),
                                            new SqlParameter("@EmailAddress", node.Element("EmailAddress").Value),
                                            new SqlParameter("@Hourly", node.Element("Hourly").Value),
                                            new SqlParameter("@HourlyDesc", node.Element("HourlyDesc").Value),
                                            new SqlParameter("@Salary", node.Element("Salary").Value),
                                            new SqlParameter("@PayRate", node.Element("PayRate").Value),
                                            new SqlParameter("@BirthDate", node.Element("BirthDate").Value),
                                            new SqlParameter("@HireDate", node.Element("HireDate").Value),
                                            new SqlParameter("@ReHireDate", node.Element("ReHireDate").Value),
                                            new SqlParameter("@AnnivDate", node.Element("AnnivDate").Value),
                                            new SqlParameter("@TermDate", node.Element("TermDate").Value),
                                            new SqlParameter("@Terminated", node.Element("Terminated").Value),
                                            new SqlParameter("@ReasonCode", node.Element("ReasonCode").Value),
                                            new SqlParameter("@AdjHireDate", node.Element("AdjHireDate").Value),
                                            new SqlParameter("@SeniorityDate", node.Element("SeniorityDate").Value),
                                            new SqlParameter("@CanRoeDate", node.Element("CanRoeDate").Value),
                                            new SqlParameter("@TermComments", node.Element("TermComments").Value),
                                            new SqlParameter("@Sex", node.Element("Sex").Value),
                                            new SqlParameter("@SexDesc", node.Element("SexDesc").Value),
                                            new SqlParameter("@EeocClass", node.Element("EeocClass").Value),
                                            new SqlParameter("@EeocClassDesc", node.Element("EeocClassDesc").Value),
                                            new SqlParameter("@Handicapped", node.Element("Handicapped").Value),
                                            new SqlParameter("@OffPhone", node.Element("OffPhone").Value),
                                            new SqlParameter("@OffPhoneExt", node.Element("OffPhoneExt").Value),
                                            new SqlParameter("@MaritalStat", node.Element("MaritalStat").Value),
                                            new SqlParameter("@MaritalStatDesc", node.Element("MaritalStatDesc").Value),
                                            new SqlParameter("@Exempt", node.Element("Exempt").Value),
                                            new SqlParameter("@PayrollCode", node.Element("PayrollCode").Value),
                                            new SqlParameter("@UnionCode", node.Element("UnionCode").Value),
                                            new SqlParameter("@TaxArea", node.Element("TaxArea").Value),
                                            new SqlParameter("@ShiftNo", node.Element("ShiftNo").Value),
                                            new SqlParameter("@RecGroup", node.Element("RecGroup").Value),
                                            new SqlParameter("@ExpenseMain", node.Element("ExpenseMain").Value),
                                            new SqlParameter("@PcheckMsg", node.Element("PcheckMsg").Value),
                                            new SqlParameter("@PcheckMsgExp", node.Element("PcheckMsgExp").Value),
                                            new SqlParameter("@BenPackage", node.Element("BenPackage").Value),
                                            new SqlParameter("@Cobra", node.Element("Cobra").Value),
                                            new SqlParameter("@CobraComplete", node.Element("CobraComplete").Value),
                                            new SqlParameter("@CobraSrcEmp", node.Element("CobraSrcEmp").Value),
                                            new SqlParameter("@CobraSrcDepId", node.Element("CobraSrcDepId").Value),
                                            new SqlParameter("@CobraEvent", node.Element("CobraEvent").Value),
                                            new SqlParameter("@CobraEventDate", node.Element("CobraEventDate").Value),
                                            new SqlParameter("@CobraEeNoteDate", node.Element("CobraEeNoteDate").Value),
                                            new SqlParameter("@CobraRightsDate", node.Element("CobraRightsDate").Value),
                                            new SqlParameter("@CobraElectCode", node.Element("CobraElectCode").Value),
                                            new SqlParameter("@CobraElectCodeDesc", node.Element("CobraElectCodeDesc").Value),
                                            new SqlParameter("@CobraElectDate", node.Element("CobraElectDate").Value),
                                            new SqlParameter("@CobraLatePayNote", node.Element("CobraLatePayNote").Value),
                                            new SqlParameter("@CobraTermLetDate", node.Element("CobraTermLetDate").Value),
                                            new SqlParameter("@CobraTermCode", node.Element("CobraTermCode").Value),
                                            new SqlParameter("@CobraTermCodeDesc", node.Element("CobraTermCodeDesc").Value),
                                            new SqlParameter("@CobraTermDate", node.Element("CobraTermDate").Value),
                                            new SqlParameter("@CobraFirstPeriod", node.Element("CobraFirstPeriod").Value),
                                            new SqlParameter("@CobraLastPeriod", node.Element("CobraLastPeriod").Value),
                                            new SqlParameter("@CobraStartDate", node.Element("CobraStartDate").Value),
                                            new SqlParameter("@CobraEndDate", node.Element("CobraEndDate").Value),
                                            new SqlParameter("@W2Flag", node.Element("W2Flag").Value),
                                            new SqlParameter("@W2MiscValue", node.Element("W2MiscValue").Value),
                                            new SqlParameter("@Veteran", node.Element("Veteran").Value),
                                            new SqlParameter("@VeteranDesc", node.Element("VeteranDesc").Value),
                                            new SqlParameter("@Rank", node.Element("Rank").Value),
                                            new SqlParameter("@Branch", node.Element("Branch").Value),
                                            new SqlParameter("@DateDischarged", node.Element("DateDischarged").Value),
                                            new SqlParameter("@DischargeType", node.Element("DischargeType").Value),
                                            new SqlParameter("@DischargeTypeDesc", node.Element("DischargeTypeDesc").Value),
                                            new SqlParameter("@CurrentService", node.Element("CurrentService").Value),
                                            new SqlParameter("@CurrentServiceDesc", node.Element("CurrentServiceDesc").Value),
                                            new SqlParameter("@PassportNo", node.Element("PassportNo").Value),
                                            new SqlParameter("@PassportExpDate", node.Element("PassportExpDate").Value),
                                            new SqlParameter("@CountryCitizen", node.Element("CountryCitizen").Value),
                                            new SqlParameter("@BirthPlace", node.Element("BirthPlace").Value),
                                            new SqlParameter("@VerifyType", node.Element("VerifyType").Value),
                                            new SqlParameter("@VerifyDate", node.Element("VerifyDate").Value),
                                            new SqlParameter("@ExpireDate", node.Element("ExpireDate").Value),
                                            new SqlParameter("@AlienNo", node.Element("AlienNo").Value),
                                            new SqlParameter("@AdmissionNo", node.Element("AdmissionNo").Value),
                                            new SqlParameter("@ListAType", node.Element("ListAType").Value),
                                            new SqlParameter("@ListADocNo", node.Element("ListADocNo").Value),
                                            new SqlParameter("@ListAExpDate", node.Element("ListAExpDate").Value),
                                            new SqlParameter("@ListBType", node.Element("ListBType").Value),
                                            new SqlParameter("@ListBDocNo", node.Element("ListBDocNo").Value),
                                            new SqlParameter("@ListBExpDate", node.Element("ListBExpDate").Value),
                                            new SqlParameter("@ListBState", node.Element("ListBState").Value),
                                            new SqlParameter("@ListBOther", node.Element("ListBOther").Value),
                                            new SqlParameter("@ListCType", node.Element("ListCType").Value),
                                            new SqlParameter("@ListCDocNo", node.Element("ListCDocNo").Value),
                                            new SqlParameter("@ListCExpDate", node.Element("ListCExpDate").Value),
                                            new SqlParameter("@ListCInsForm", node.Element("ListCInsForm").Value),
                                            new SqlParameter("@ApprovalDate", node.Element("ApprovalDate").Value),
                                            new SqlParameter("@ApprovedBy", node.Element("ApprovedBy").Value),
                                            new SqlParameter("@ApprovedByTitle", node.Element("ApprovedByTitle").Value),
                                            new SqlParameter("@ReviewDate", node.Element("ReviewDate").Value),
                                            new SqlParameter("@EmergenyContact", node.Element("EmergenyContact").Value),
                                            new SqlParameter("@EmergAddress1", node.Element("EmergAddress1").Value),
                                            new SqlParameter("@EmergAddress2", node.Element("EmergAddress2").Value),
                                            new SqlParameter("@EmergCity", node.Element("EmergCity").Value),
                                            new SqlParameter("@EmergState", node.Element("EmergState").Value),
                                            new SqlParameter("@EmergZipCode", node.Element("EmergZipCode").Value),
                                            new SqlParameter("@EmergRelation", node.Element("EmergRelation").Value),
                                            new SqlParameter("@EmergPrimPhoneNo", node.Element("EmergPrimPhoneNo").Value),
                                            new SqlParameter("@EmergSecPhoneNo", node.Element("EmergSecPhoneNo").Value),
                                            new SqlParameter("@Physician", node.Element("Physician").Value),
                                            new SqlParameter("@PhysicianPhoneNo", node.Element("PhysicianPhoneNo").Value),
                                            new SqlParameter("@NextPhysical", node.Element("NextPhysical").Value),
                                            new SqlParameter("@LastPhysical", node.Element("LastPhysical").Value),
                                            new SqlParameter("@PhyRenewMonth", node.Element("PhyRenewMonth").Value),
                                            new SqlParameter("@DonationDate", node.Element("DonationDate").Value),
                                            new SqlParameter("@PhysicalResult", node.Element("PhysicalResult").Value),
                                            new SqlParameter("@WorkRest", node.Element("WorkRest").Value),
                                            new SqlParameter("@MedicalCode", node.Element("MedicalCode").Value),
                                            new SqlParameter("@Height", node.Element("Height").Value),
                                            new SqlParameter("@Weight", node.Element("Weight").Value),
                                            new SqlParameter("@BloodType", node.Element("BloodType").Value),
                                            new SqlParameter("@RhFactor", node.Element("RhFactor").Value),
                                            new SqlParameter("@NoFedEx", node.Element("NoFedEx").Value),
                                            new SqlParameter("@NoStateEx", node.Element("NoStateEx").Value),
                                            new SqlParameter("@NoLocalEx", ""),
                                            new SqlParameter("@FedFilingStat", node.Element("FedFilingStat").Value),
                                            new SqlParameter("@StateFilingStat", node.Element("StateFilingStat").Value),
                                            new SqlParameter("@LocalFilingStat", node.Element("SecondLocalFilingStatus").Value),
                                            new SqlParameter("@AllowReimb", node.Element("AllowReimb").Value),
                                            new SqlParameter("@NoOfFtes", node.Element("NoOfFtes").Value),
                                            new SqlParameter("@FteClass", node.Element("FteClass").Value),
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
                                            new SqlParameter("@MiscAlpha5", node.Element("MiscAlpha5").Value),
                                            new SqlParameter("@MiscAlpha6", node.Element("MiscAlpha6").Value),
                                            new SqlParameter("@MiscAlpha7", node.Element("MiscAlpha7").Value),
                                            new SqlParameter("@MiscAlpha8", node.Element("MiscAlpha8").Value),
                                            new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                            new SqlParameter("@UserAlpha2", node.Element("UserAlpha2").Value),
                                            new SqlParameter("@UserAlpha3", node.Element("UserAlpha3").Value),
                                            new SqlParameter("@UserAlpha4", node.Element("UserAlpha4").Value),
                                            new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                            new SqlParameter("@MiscDate2", node.Element("MiscDate2").Value),
                                            new SqlParameter("@UserDate1", node.Element("UserDate1").Value),
                                            new SqlParameter("@ListATypeDesc", node.Element("ListATypeDesc").Value),
                                            new SqlParameter("@Supervisor", node.Element("SupervisorId").Value),
                                            new SqlParameter("@ListCTypeDesc", node.Element("ListCTypeDesc").Value),
                                            new SqlParameter("@UITaxArea", node.Element("UiTaxArea").Value),
                                           // new SqlParameter("@ReHireDate", node.Element("ReHireDate").Value),
                                            new SqlParameter("@VerificationTypeDesc", node.Element("VerificationTypeDesc").Value),
                                            new SqlParameter("@EmployeeFTEs", node.Element("EmployeeFtes").Value),
                                            new SqlParameter("@AllowPaychecksAfterTerm", node.Element("AllowPaychecksAfterTerm").Value),
                                            new SqlParameter("@ListBTypeDesc", node.Element("ListBTypeDesc").Value),
                                            new SqlParameter("@ListAOther", node.Element("ListAOther").Value),
                                            new SqlParameter("@SortField1", node.Element("SortField1").Value),
                                            new SqlParameter("@SortField2", node.Element("SortField2").Value),
                                            new SqlParameter("@SortField3", node.Element("SortField3").Value),
                                            new SqlParameter("@SortField4", node.Element("SortField4").Value),
                                            new SqlParameter("@SortField5", node.Element("SortField5").Value),
                                            new SqlParameter("@AllowDepartmentTransfer", node.Element("AllowDepartmentTransfer").Value),
                                            new SqlParameter("@AllowJobTransfer", node.Element("AllowJobTransfer").Value),
                                            new SqlParameter("@AllowPunchAcrossDays", node.Element("AllowPunchAcrossDays").Value),
                                            new SqlParameter("@ListOfValidDepartments", node.Element("ListOfValidDepartments").Value),
                                            new SqlParameter("@ListofValidJobs", node.Element("ListofValidJobs").Value),
                                            new SqlParameter("@SupervisorId", node.Element("SupervisorId").Value),
                                            new SqlParameter("@WorkEmailAddress", node.Element("WorkEmailAddress").Value),
                                            new SqlParameter("@PseudoUserID", node.Element("PseudoUserID").Value),
                                            new SqlParameter("@TimeKeepingProfile", node.Element("TimeKeepingProfile").Value),
                                            new SqlParameter("@ScheduleID", node.Element("ScheduleID").Value),
                                            new SqlParameter("@SchedulingMethod", node.Element("SchedulingMethod").Value),
                                            new SqlParameter("@Team", node.Element("Team").Value),
                                            new SqlParameter("@SupervisoryEE", node.Element("SupervisoryEE").Value),
                                            new SqlParameter("@PsGroupNo", node.Element("PsGroupNo").Value),
                                            new SqlParameter("@LastEvalDateAndMiltime", node.Element("LastEvalDateAndMiltime").Value),
                                            new SqlParameter("@CellPhone", node.Element("CellPhone").Value),
                                            new SqlParameter("@NameSuffix", node.Element("NameSuffix").Value),
                                            new SqlParameter("@OverrideUIDeductCode", node.Element("OverrideUIDeductCode").Value),
                                            new SqlParameter("@NewHireBenefitsProcessed", node.Element("NewHireBenefitsProcessed").Value),
                                            new SqlParameter("@Eligible", node.Element("Eligible").Value),
                                            new SqlParameter("@GeographicCode", node.Element("GeographicCode").Value),
                                            new SqlParameter("@Rehire", node.Element("Rehire").Value),
                                            new SqlParameter("@MiscDec3", node.Element("MiscDec3").Value),
                                            new SqlParameter("@MiscDec4", node.Element("MiscDec4").Value),
                                            new SqlParameter("@MiscLog3", node.Element("MiscLog3").Value),
                                            new SqlParameter("@MiscLog4", node.Element("MiscLog4").Value),
                                            new SqlParameter("@MiscAlpha9", node.Element("MiscAlpha9").Value),
                                            new SqlParameter("@MiscAlpha10", node.Element("MiscAlpha10").Value),
                                            new SqlParameter("@MiscAlpha11", node.Element("MiscAlpha11").Value),
                                            new SqlParameter("@MiscAlpha12", node.Element("MiscAlpha12").Value),
                                            new SqlParameter("@MiscAlpha13", node.Element("MiscAlpha13").Value),
                                            new SqlParameter("@MiscAlpha14", node.Element("MiscAlpha14").Value),
                                            new SqlParameter("@MiscDate3", node.Element("MiscDate3").Value),
                                            new SqlParameter("@EESpecificTimeEntryTemplate", node.Element("EESpecificTimeEntryTemplate").Value),
                                            new SqlParameter("@TimeEntryTemplate", node.Element("TimeEntryTemplate").Value));
                                    break;
                                case "mainacct":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Mainacct",
                                             new SqlParameter("@loads_id", loadsId),
                                             new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                             new SqlParameter("@MainAcct", node.Element("MainAcct").Value),
                                             new SqlParameter("@MajorAcct", node.Element("MajorAcct").Value),
                                             new SqlParameter("@MinorAcct", node.Element("MinorAcct").Value),
                                             new SqlParameter("@Description", node.Element("Description").Value),
                                             new SqlParameter("@ShortDesc", node.Element("ShortDesc").Value),
                                             new SqlParameter("@Alpha", node.Element("Alpha").Value),
                                             new SqlParameter("@LevOfDet", node.Element("LevOfDet").Value),
                                             new SqlParameter("@LevOfDetDesc", node.Element("LevOfDetDesc").Value),
                                             new SqlParameter("@AllocID", node.Element("AllocId").Value),
                                             new SqlParameter("@AcctCat", node.Element("AcctCat").Value),
                                             new SqlParameter("@AcctCatDesc", node.Element("AcctCatDesc").Value),
                                             new SqlParameter("@Function", node.Element("Function").Value),
                                             new SqlParameter("@AcctClass", node.Element("AcctClass").Value),
                                             new SqlParameter("@ProjectExc", node.Element("ProjectExc").Value),
                                             new SqlParameter("@ProjectExcDesc", node.Element("ProjectExcDesc").Value),
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
                                             new SqlParameter("@ExcludedJournalCodes", node.Element("ExcludedJournalCodes").Value),
                                             new SqlParameter("@AccountType", node.Element("AccountType").Value),
                                             new SqlParameter("@SearchField", node.Element("SearchField").Value),
                                             new SqlParameter("@SectionId", node.Element("SectionId").Value),
                                             new SqlParameter("@RequiredProjectType", node.Element("RequiredProjectType").Value),
                                             new SqlParameter("@MainacctKey", node.Element("MainacctKey").Value),
                                             new SqlParameter("@UpdateSeqNo", node.Element("UpdateSeqNo").Value),
                                             new SqlParameter("@ColsolMainAcct", node.Element("ConsolMainAcct").Value),
                                             new SqlParameter("@RequiredJobType", node.Element("RequiredJobType").Value),
                                             new SqlParameter("@DateTimeChanged", ""));
                                    break;
                                case "pperiod":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Pperiod",
                                        new SqlParameter("@loads_id", loadsId),
                                        new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                        new SqlParameter("@PayrollCode", node.Element("PayrollCode").Value),
                                        new SqlParameter("@PayPeriod", node.Element("PayPeriod").Value),
                                        new SqlParameter("@StartDate", node.Element("StartDate").Value),
                                        new SqlParameter("@EndDate", node.Element("EndDate").Value),
                                        new SqlParameter("@PperiodId", node.Element("PperiodId").Value),
                                        new SqlParameter("@Description", node.Element("Description").Value),
                                        new SqlParameter("@PeriodStatus", node.Element("PeriodStatus").Value),
                                        new SqlParameter("@PeriodStatusDesc", node.Element("PeriodStatusDesc").Value),
                                        new SqlParameter("@DaysPerPeriod", node.Element("DaysPerPeriod").Value),
                                        new SqlParameter("@HrsPerPeriod", node.Element("HrsPerPeriod").Value),
                                        new SqlParameter("@PcheckMsg", node.Element("PcheckMsg").Value),
                                        new SqlParameter("@PostMo", node.Element("PostMo").Value),
                                        new SqlParameter("@PostYr", node.Element("PostYr").Value),
                                        new SqlParameter("@CalPeriod", node.Element("CalPeriod").Value),
                                        new SqlParameter("@PctAllocNext", node.Element("PctAllocNext").Value),
                                        new SqlParameter("@CheckDate", node.Element("CheckDate").Value),
                                        new SqlParameter("@CalYr", node.Element("CalYr").Value),
                                        new SqlParameter("@PayPeriodNo", node.Element("PayPeriodNo").Value),
                                        new SqlParameter("@GrossAmt", node.Element("GrossAmt").Value),
                                        new SqlParameter("@EeDeductAmt", node.Element("EeDeductAmt").Value),
                                        new SqlParameter("@ErDeductAmt", node.Element("ErDeductAmt").Value),
                                        new SqlParameter("@NetAmt", node.Element("NetAmt").Value),
                                        new SqlParameter("@Hrs", node.Element("Hrs").Value),
                                        new SqlParameter("@NoOfChecks", node.Element("NoOfChecks").Value),
                                        new SqlParameter("@CnMinHrsPp", node.Element("CnMinHrsPp").Value),
                                        new SqlParameter("@CnMinEarnPp", node.Element("CnMinEarnPp").Value),
                                        new SqlParameter("@CnMaxEarnPp", node.Element("CnMaxEarnPp").Value),
                                        new SqlParameter("@CnMaxPremiumPp", node.Element("CnMaxPremiumPp").Value),
                                        new SqlParameter("@CnMaxPremiumYr", node.Element("CnMaxPremiumYr").Value),
                                        new SqlParameter("@CnMaxEarnYr", node.Element("CnMaxEarnYr").Value),
                                        new SqlParameter("@CnMaxPremium53", node.Element("CnMaxPremium53").Value),
                                        new SqlParameter("@CnMaxEarn53", node.Element("CnMaxEarn53").Value),
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
                                        new SqlParameter("@PayPeriodSequence", node.Element("PayPeriodSequence").Value),
                                        new SqlParameter("@CheckMonth", node.Element("CheckMonth").Value),
                                        new SqlParameter("@Processed", node.Element("Processed").Value),
                                        new SqlParameter("@SortField1", node.Element("SortField1").Value),
                                        new SqlParameter("@SortField2", node.Element("SortField2").Value),
                                        new SqlParameter("@SortField3", node.Element("SortField3").Value),
                                        new SqlParameter("@SortField4", node.Element("SortField4").Value),
                                        new SqlParameter("@SortField5", node.Element("SortField5").Value),
                                        new SqlParameter("@TimeAttendStatus", node.Element("TimeAttendStatus").Value));
                                    break;
                                case "pyjobcd":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Pyjobcd",
                                           new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                            new SqlParameter("@JobCode", node.Element("JobCode").Value),
                                            new SqlParameter("@ExpenseMain", node.Element("ExpenseMain").Value),
                                            new SqlParameter("@EeocCat", node.Element("EeocCat").Value),
                                            new SqlParameter("@EeocCatDesc", node.Element("EeocCatDesc").Value),
                                            new SqlParameter("@Supervisory", node.Element("Supervisory").Value),
                                            new SqlParameter("@Description", node.Element("Description").Value),
                                            new SqlParameter("@ShortDesc", node.Element("ShortDesc").Value),
                                            new SqlParameter("@MinSalary", node.Element("MinSalary").Value),
                                            new SqlParameter("@MidSalary", node.Element("MidSalary").Value),
                                            new SqlParameter("@MaxSalary", node.Element("MaxSalary").Value),
                                            new SqlParameter("@EvalDate", node.Element("EvalDate").Value),
                                            new SqlParameter("@CreateDate", node.Element("CreateDate").Value),
                                            new SqlParameter("@SlotKnowHow", node.Element("SlotKnowHow").Value),
                                            new SqlParameter("@SlotAcctability", node.Element("SlotAcctability").Value),
                                            new SqlParameter("@SlotProblemSolve", node.Element("SlotProblemSolve").Value),
                                            new SqlParameter("@SlotWorkCond", node.Element("SlotWorkCond").Value),
                                            new SqlParameter("@PointsKnowHow", node.Element("PointsKnowHow").Value),
                                            new SqlParameter("@PointsAcctability", node.Element("PointsAcctability").Value),
                                            new SqlParameter("@PointsProblemSolve", node.Element("PointsProblemSolve").Value),
                                            new SqlParameter("@PointsWorkCond", node.Element("PointsWorkCond").Value),
                                            new SqlParameter("@ProblemSolvePct", node.Element("ProblemSolvePct").Value),
                                            new SqlParameter("@TotalPoints", node.Element("TotalPoints").Value),
                                            new SqlParameter("@Multiplier", node.Element("Multiplier").Value),
                                            new SqlParameter("@BaseAmount", node.Element("BaseAmount").Value),
                                            new SqlParameter("@MinimumFactor", node.Element("MinimumFactor").Value),
                                            new SqlParameter("@MaximumFactor", node.Element("MaximumFactor").Value),
                                            new SqlParameter("@OvertimeFactor", node.Element("OvertimeFactor").Value),
                                            new SqlParameter("@HayMinSalary", node.Element("HayMinSalary").Value),
                                            new SqlParameter("@HayMidSalary", node.Element("HayMidSalary").Value),
                                            new SqlParameter("@HayMaxSalary", node.Element("HayMaxSalary").Value),
                                            new SqlParameter("@HayEvalDate", node.Element("HayEvalDate").Value),
                                            new SqlParameter("@ShiftDiffTypes", node.Element("ShiftDiffTypes").Value),
                                            new SqlParameter("@ShiftDiffTypesDesc", node.Element("ShiftDiffTypesDesc").Value),
                                            new SqlParameter("@ShiftDiffs", node.Element("ShiftDiffs").Value),
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
                                            new SqlParameter("@AnnualHrs", node.Element("AnnualHrs").Value),
                                            new SqlParameter("@ListOfXferJobCodes", node.Element("ListOfXferJobCodes").Value),
                                            new SqlParameter("@SOCcode", node.Element("SOCcode").Value),
                                            new SqlParameter("@DfltWcCat", node.Element("DfltWcCat").Value));
                                    break;
                                case "pyunion":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Pyunion",
                                            new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                            new SqlParameter("@UnionCode", node.Element("UnionCode").Value),
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
                                            new SqlParameter("@UserInt1", node.Element("UserInt1").Value),
                                            new SqlParameter("@MiscDec1", node.Element("MiscDec1").Value),
                                            new SqlParameter("@UserDec1", node.Element("UserDec1").Value),
                                            new SqlParameter("@MiscLog1", node.Element("MiscLog1").Value),
                                            new SqlParameter("@UserLog1", node.Element("UserLog1").Value),
                                            new SqlParameter("@MiscAlpha1", node.Element("MiscAlpha1").Value),
                                            new SqlParameter("@MiscAlpha2", node.Element("MiscAlpha2").Value),
                                            new SqlParameter("@UserAlpha1", node.Element("UserAlpha1").Value),
                                            new SqlParameter("@MiscDate1", node.Element("MiscDate1").Value),
                                            new SqlParameter("@UserDate1", node.Element("UserDate1").Value));
                                    break;
                                case "tcard2":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Tcard2",
                                            new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@OpId", node.Element("OpId").Value),
                                            new SqlParameter("@ControlNo", node.Element("ControlNo").Value),
                                            new SqlParameter("@TrxNo", node.Element("TrxNo").Value),
                                            new SqlParameter("@LineNo", node.Element("LineNo").Value),
                                            new SqlParameter("@CheckDocId", node.Element("CheckDocId").Value),
                                            new SqlParameter("@EmployeeNo", node.Element("EmployeeNo").Value),
                                            new SqlParameter("@Project", node.Element("Project").Value),
                                            new SqlParameter("@ProjectWb", node.Element("ProjectWb").Value),
                                            new SqlParameter("@EarningsCode", node.Element("EarningsCode").Value),
                                            new SqlParameter("@JobCode", node.Element("JobCode").Value),
                                            new SqlParameter("@BusUnit", node.Element("BusUnit").Value),
                                            new SqlParameter("@ExpenseAcct", node.Element("ExpenseAcct").Value),
                                            new SqlParameter("@DiffExpAcct", node.Element("DiffExpAcct").Value),
                                            new SqlParameter("@TaxArea", node.Element("TaxArea").Value),
                                            new SqlParameter("@StateCode", node.Element("StateCode").Value),
                                            new SqlParameter("@AllocId", node.Element("AllocId").Value),
                                            new SqlParameter("@Description", node.Element("Description").Value),
                                            new SqlParameter("@TrxDate", node.Element("TrxDate").Value),
                                            new SqlParameter("@Hrs", node.Element("Hrs").Value),
                                            new SqlParameter("@PayRate", node.Element("PayRate").Value),
                                            new SqlParameter("@Amt", node.Element("Amt").Value),
                                            new SqlParameter("@DiffAmt", node.Element("DiffAmt").Value),
                                            new SqlParameter("@ShiftNo", node.Element("ShiftNo").Value),
                                            new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                            new SqlParameter("@RecTimeId", node.Element("RecTimeId").Value),
                                            new SqlParameter("@LineId", node.Element("LineId").Value),
                                            new SqlParameter("@EffectDate", node.Element("EffectDate").Value),
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
                                            new SqlParameter("@SearchField", node.Element("SearchField").Value),
                                            new SqlParameter("@TimeAdded", node.Element("TimeAdded").Value),
                                            new SqlParameter("@TimeAddedDesc", node.Element("TimeAddedDesc").Value),
                                            new SqlParameter("@RPExpenseAcct", node.Element("RpExpenseacct").Value));
                                    break;
                                case "trx1099":
                                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Trx1099",
                                            new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                            new SqlParameter("@TrxNo", node.Element("TrxNo").Value),
                                            new SqlParameter("@Vendor", node.Element("Vendor").Value),
                                            new SqlParameter("@TypeOf1099", node.Element("TypeOf1099").Value),
                                            new SqlParameter("@CalYr", node.Element("CalYr").Value),
                                            new SqlParameter("@BoxNo", node.Element("BoxNo").Value),
                                            new SqlParameter("@Amt", node.Element("Amt").Value),
                                            new SqlParameter("@TextValue", node.Element("TextValue").Value),
                                            new SqlParameter("@CheckNo", node.Element("CheckNo").Value),
                                            new SqlParameter("@CheckOpId", node.Element("CheckOpId").Value),
                                            new SqlParameter("@CheckControlNo", node.Element("CheckControlNo").Value),
                                            new SqlParameter("@CheckTrxNo", node.Element("CheckTrxNo").Value),
                                            new SqlParameter("@Description", node.Element("Description").Value),
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
                                case "vendor":
                                   ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Vendor",
                                            new SqlParameter("@loads_id", loadsId),
                                            new SqlParameter("@CompanyId", node.Element("CompanyId").Value),
                                            new SqlParameter("@Vendor", node.Element("Vendor").Value),
                                            new SqlParameter("@Alpha", node.Element("Alpha").Value),
                                            new SqlParameter("@SepCheck", node.Element("SepCheck").Value),
                                            new SqlParameter("@MiscVendor", node.Element("MiscVendor").Value),
                                            new SqlParameter("@Terminated", node.Element("Terminated").Value),
                                            new SqlParameter("@Termination", node.Element("Termination").Value),
                                            new SqlParameter("@GenDist", node.Element("GenDist").Value),
                                            new SqlParameter("@PayStatus", node.Element("PayStatus").Value),
                                            new SqlParameter("@PayStatusDesc", node.Element("PayStatusDesc").Value),
                                            new SqlParameter("@BusUnit", node.Element("BusUnit").Value),
                                            new SqlParameter("@VendorType", node.Element("VendorType").Value),
                                            new SqlParameter("@TermsCode", node.Element("TermsCode").Value),
                                            new SqlParameter("@ApAcct", node.Element("ApAcct").Value),
                                            new SqlParameter("@ExpenseAcct", node.Element("ExpenseAcct").Value),
                                            new SqlParameter("@TypeOf1099", node.Element("TypeOf1099").Value),
                                            new SqlParameter("@RemitTo", node.Element("RemitTo").Value),
                                            new SqlParameter("@SendPoTo", node.Element("SendPoTo").Value),
                                            new SqlParameter("@Buyer", node.Element("Buyer").Value),
                                            new SqlParameter("@TaxAuth", node.Element("TaxAuth").Value),
                                            new SqlParameter("@Template", node.Element("Template").Value),
                                            new SqlParameter("@DutyCode", node.Element("DutyCode").Value),
                                            new SqlParameter("@CurrCode", node.Element("CurrCode").Value),
                                            new SqlParameter("@Hold", node.Element("Hold").Value),
                                            new SqlParameter("@HoldThrough", node.Element("HoldThrough").Value),
                                            new SqlParameter("@HoldNote", node.Element("HoldNote").Value),
                                            new SqlParameter("@CreditLimit", node.Element("CreditLimit").Value),
                                            new SqlParameter("@Balance", node.Element("Balance").Value),
                                            new SqlParameter("@BalOnPo", node.Element("BalOnPo").Value),
                                            new SqlParameter("@DisputedAmt", node.Element("DisputedAmt").Value),
                                            new SqlParameter("@AmtOnHold", node.Element("AmtOnHold").Value),
                                            new SqlParameter("@MinimumPo", node.Element("MinimumPo").Value),
                                            new SqlParameter("@ReviewCycle", node.Element("ReviewCycle").Value),
                                            new SqlParameter("@LastReview", node.Element("LastReview").Value),
                                            new SqlParameter("@NextReview", node.Element("NextReview").Value),
                                            new SqlParameter("@Name", node.Element("Name").Value),
                                            new SqlParameter("@Address1", node.Element("Address1").Value),
                                            new SqlParameter("@Address2", node.Element("Address2").Value),
                                            new SqlParameter("@City", node.Element("City").Value),
                                            new SqlParameter("@State", node.Element("State").Value),
                                            new SqlParameter("@ZipCode", node.Element("ZipCode").Value),
                                            new SqlParameter("@Country", node.Element("Country").Value),
                                            new SqlParameter("@Phone", node.Element("Phone").Value),
                                            new SqlParameter("@FaxNo", node.Element("FaxNo").Value),
                                            new SqlParameter("@EmailAddress", node.Element("EmailAddress").Value),
                                            new SqlParameter("@FirstTrx", node.Element("FirstTrx").Value),
                                            new SqlParameter("@LastTrx", node.Element("LastTrx").Value),
                                            new SqlParameter("@LastPayment", node.Element("LastPayment").Value),
                                            new SqlParameter("@PartialUnits", node.Element("PartialUnits").Value),
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
                                            new SqlParameter("@VendorAccount", node.Element("VendorAccount").Value),
                                            new SqlParameter("@VendorBankTransit", node.Element("VendorBankTransit").Value),
                                            new SqlParameter("@VendorBankAcctNo", node.Element("VendorBankAcctNo").Value),
                                            new SqlParameter("@UseTaxState", node.Element("PrenoteStatus").Value),
                                            new SqlParameter("@PrenoteStatus", node.Element("PrenoteStatus").Value),
                                            new SqlParameter("@BankAcctType", node.Element("BankAcctType").Value),
                                            new SqlParameter("@DefaultBoxNo", node.Element("DefaultBoxNo").Value),
                                            new SqlParameter("@PrenoteDate", node.Element("PrenoteDate").Value),
                                            new SqlParameter("@Garnished", node.Element("Garnished").Value),
                                            new SqlParameter("@GarnRemitTo", node.Element("GarnRemitTo").Value),
                                            new SqlParameter("@GarnCaseID", node.Element("GarnCaseId").Value),
                                            new SqlParameter("@GarnSocSecNo", node.Element("GarnSocSecNo").Value),
                                            new SqlParameter("@GarnFirstName", node.Element("GarnFirstName").Value),
                                            new SqlParameter("@GarnLastName", node.Element("GarnLastName").Value),
                                            new SqlParameter("@GarnFIPSCode", node.Element("GarnFipsCode").Value),
                                            new SqlParameter("@GarnEffectDate", node.Element("GarnEffectDate").Value),
                                            new SqlParameter("@GarnExpireDate", node.Element("GarnExpireDate").Value),
                                            new SqlParameter("@GarnMethod", node.Element("GarnMethod").Value),
                                            new SqlParameter("@GarnPct", node.Element("GarnPct").Value),
                                            new SqlParameter("@GarnAmt", node.Element("GarnAmt").Value),
                                            new SqlParameter("@GarnEFT", node.Element("GarnEft").Value),
                                            new SqlParameter("@Address3", node.Element("Address3").Value),
                                            new SqlParameter("@SuspendAchThru", node.Element("SuspendAchThru").Value),
                                            new SqlParameter("@NewVendorID", node.Element("NewVendorID").Value),
                                            new SqlParameter("@DefaultTrxDescription", node.Element("DefaultTrxDescription").Value),
                                            new SqlParameter("@BirthDate", node.Element("BirthDate").Value),
                                            new SqlParameter("@SpecialHandling", node.Element("SpecialHandling").Value),
                                            new SqlParameter("@SpecialHandlingSortCode", node.Element("SpecialHandlingSortCode").Value),
                                            new SqlParameter("@DefaultCreditCardVendor", node.Element("DefaultCreditCardVendor").Value),
                                            new SqlParameter("@DunsNumber", node.Element("DunsNumber").Value),
                                            new SqlParameter("@Shipper", node.Element("Shipper").Value),
                                            new SqlParameter("@NextInsurDeductDate", node.Element("NextInsurDeductDate").Value),
                                            new SqlParameter("@CarrierInsuranceAmt", node.Element("CarrierInsuranceAmt").Value),
                                            new SqlParameter("@SuppressACHReceipt", node.Element("SuppressACHReceipt").Value),
                                            new SqlParameter("@ExclForFixedAssetAdditions", node.Element("ExclForFixedAssetAdditions").Value),
                                            new SqlParameter("@SuppresRemittanceAdvice", node.Element("SuppresRemittanceAdvice").Value),
                                            new SqlParameter("@PayGroup", node.Element("PayGroup").Value),
                                            new SqlParameter("@W9Received", node.Element("W9Received").Value));
                                    break;

                            }
                        }

                    

                    WriteToJobLog(JobLogMessageType.INFO, "Completed data load");

                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Insert_Update_Loads_Latest_By_Table",
                                                    new SqlParameter("@pvchrTableName", table["table_name"].ToString()),
                                                    new SqlParameter("@pintLoadsID", loadsId));

                    WriteToJobLog(JobLogMessageType.INFO, "Deleting loads record");

                    ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Delete_Loads",
                                                    new SqlParameter("@pvchrTableName", table["table_name"].ToString()));

                    //this only runs for the Employee table
                    RetrievePostLoad(table["table_name"].ToString(), loadsId);
                    }
                }

                WriteToJobLog(JobLogMessageType.INFO, "Deleting final load record");

                ExecuteNonQuery(DatabaseConnectionStringNames.SBSReports, "Proc_Delete_Loads",
                                new SqlParameter("@pvchrTableName", ""));



            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }


        private void RetrievePostLoad(string tableName, Int64 loadsId)
        {
            WriteToJobLog(JobLogMessageType.INFO, "Checking for post load routines");

            if (tableName == "Employee")
            {
                ExecuteNonQuery(DatabaseConnectionStringNames.Trade, "Proc_Insert_Update_Users",
                                            new SqlParameter("@pvchrSBSReportsServerInstance", GetConfigurationKeyValue("RemoteServerName")),
                                            new SqlParameter("@pvchrSBSReportsDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                            new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                            new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")),
                                            new SqlParameter("@pintEmployeeLoadsID", loadsId));
            }
            else
                WriteToJobLog(JobLogMessageType.INFO, "No post load routines");
        }
    }
}
