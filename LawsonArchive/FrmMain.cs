using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using BSGlobals;

namespace LawsonArchive
{
    public partial class FrmMain : Form
    {

        #region Declarations
        // Constants
        const string JobName = "LawsonArchive";

        // Class declarations

        // Other global stuff
        ActiveDirectory UserInfo;
        VersionStatusBar StatusBar;

        PrintDocument document = new PrintDocument();
        PrintDialog printdialog = new PrintDialog();
        PrintPreviewDialog previewdialog = new PrintPreviewDialog();

        #endregion

        #region Initialization
        public FrmMain()
        {
            InitializeComponent();

            // Job log start
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job starting", JobName);

            // Get configuration values.  

            //EXAMPLE:  bool lookbackokay = int.TryParse(Config.GetConfigurationKeyValue("Purchasing", "LookbackInYears"), out LookbackInYears);

            // Create event handlers if any 
            //EXAMPLE: TxtAddressLine1.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);

            // Menu strip initialization (where needed) TBD TBD TBD - Can't get this to work but looks identical to PurchaseOrders!!!!!
            //MainMenuStrip.Renderer = new CustomMenuStripRenderer();

            // Get the current (logged-in) username.  It will be in the form DOMAIN\username
            UserInfo = new ActiveDirectory();
            bool UserOkay = UserInfo.CheckUserCredentials(new List<string> { "BSOU_LawsonReports", "BSOU_LawsonUsers", "bsadmin" });
            if (!UserOkay)
            {
                BroadcastError("You do not have the appropriate credentials (LawsonReports) to run this app.", null);
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job completed", JobName);
                System.Environment.Exit(1);
            }

            // Add status bar (2 segment default, with version)
            StatusBar = new VersionStatusBar(this);

            // Populate employee combo box, by Employee name

            PopulateEmployeeComboBox(true);

            document.PrintPage += new PrintPageEventHandler(Document_PrintPage);

        }

        void Document_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawString(TxtEmployeeData.Text + TxtWages.Text, new Font("Courier New", 10, FontStyle.Regular), Brushes.Black, 20, 20);
        }
        #endregion

        #region Data Display Functions

        #endregion

        #region Button Rendering Functions

        #endregion

        #region Safe Value Assignments

        /// <summary>
        /// A generic way to safely copy any string-able value from a dictionary into any control that has a .Text property.
        /// </summary>
        /// <param name="t"></param>
        /// <param name="dic"></param>
        /// <param name="s"></param>
        private static void SafeText(object t, Dictionary<string, object> dic, string s)
        {
            // See if the passed-in object has a .Text property
            var prop = t.GetType().GetProperty("Text", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
            try
            {
                // If the object has a .Text property then assign the selected dictionary entry value to it, and color it black.
                if (prop != null)
                {
                    prop.SetValue(t, dic[s].ToString(), null);
                    var forecolor = t.GetType().GetProperty("ForeColor", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance);
                    if (forecolor != null) forecolor.SetValue(t, Color.Black, null);
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate control from string:  " + ex.ToString(), JobName);
                prop.SetValue(t, "", null);
            }
        }

#if false
        private void SafeTextBox(TextBox t, Dictionary<string, object> dic, string s)
        {
            try
            {
                t.Text = dic[s].ToString();
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate textbox from string:  " + ex.ToString(), JobName);
                t.Text = "";
            }
        }

        private void SafeComboBox(ComboBox t, Dictionary<string, object> dic, string s)
        {
            try
            {
                t.Text = dic[s].ToString();
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate combobox from string:  " + ex.ToString(), JobName);
                t.Text = "";
            }
        }
#endif

        private void SafeDateBox(DateTimePicker t, Dictionary<string, object> dic, string s)
        {
            try
            {
                t.Value = Convert.ToDateTime(dic[s].ToString());
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate datetimepicker from string:  " + ex.ToString(), JobName);
            }
        }

        private void SafeRadioBox(RadioButton radActive, RadioButton radInactive, Dictionary<string, object> dic, string s)
        {
            try
            {
                if ((bool)dic[s])
                {
                    radActive.Checked = true;
                }
                else
                {
                    radInactive.Checked = true;
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate radio buttons from string:  " + ex.ToString(), JobName);
            }
        }

        private void SafeComboBox(ComboBox t, string s)
        {
            try
            {
                t.Text = s;
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate combobox from string:  " + ex.ToString(), JobName);
                t.Text = "";
            }
        }

        #endregion

        #region Timer-related Functions

        #endregion

        #region SQL
        public static SqlDataReader SQLQuery(string qryName)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.Purchasing,
                CommandType.StoredProcedure,
                qryName);
            return (rdr);
        }

        public static SqlDataReader SQLQuery(string qryName, SqlParameter[] orderParams)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.Purchasing,
                CommandType.StoredProcedure,
                qryName,
                orderParams);
            return (rdr);
        }

        public static void SQLProcCall(string procName, SqlParameter[] Params)
        {
            DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.Purchasing,
                CommandType.StoredProcedure,
                procName,
                Params);
        }

        /// <summary>
        /// A utility to return any value from a SQL Query (as long as the underlying SQL type is knnown apriori).
        /// USAGE:  <type T> x = (T)GetSQLValue(SQLReader, FieldName)
        /// NOTE:   Declarations that use this function MUST be declared nullable!  (i.e., int? var1, double? var2, etc.
        /// </summary>
        /// <param name="rdr"></param>
        /// <param name="s"></param>
        /// <returns>(T)Value</returns>
        public static object GetSQLValue(SqlDataReader rdr, string s)
        {
            // A utility to return any value from a SQL Query (as long as the underlying SQL type is knnown apriori).
            // USAGE:
            //   <type T> x = (T)GetSQLValue(SQLReader, FieldName)

            // Because SQL can return a dbnull, there is no way to determine the actual value type.  
            //   This requires that all declarations must be nullable.
            if (rdr[s] != null)
            {
                string t = rdr[s].GetType().ToString().ToLower();
                switch (t)
                {
                    case "system.string":
                        try { return rdr[s].ToString(); } catch { return ""; }
                    case "system.int32":
                        try
                        {
                            string i = rdr[s].ToString();
                            bool iokay = int.TryParse(i, out int ii);
                            return (iokay ? ii : 0);
                        }
                        catch { return 0; }
                    case "system.decimal":
                        try
                        {
                            string d = rdr[s].ToString();
                            bool dokay = Double.TryParse(d, out double dd);
                            return (dokay ? dd : 0);
                        }
                        catch { return 0; }
                    case "system.float":
                        try
                        {
                            string f = rdr[s].ToString();
                            bool fokay = float.TryParse(f, out float ff);
                            return (fokay ? ff : 0);
                        }
                        catch { return 0; }
                    case "system.bit":
                        try
                        {
                            string b = rdr[s].ToString();
                            bool iokay = int.TryParse(b, out int bb);
                            return (iokay ? (bb == 0 ? false : true) : false);
                        }
                        catch { return false; }
                    case "system.bool":
                        try
                        {
                            string b = rdr[s].ToString();
                            bool iokay = int.TryParse(b, out int bb);
                            return (iokay ? (bb == 0 ? false : true) : false);
                        }
                        catch { return false; }
                    case "system.dbnull":
                        // Because SQL can return a dbnull, there is no way to determine the actual value type.  
                        //   This requires that all declarations must be nullable.
                        return null;
                    default:
                        // TBD Check other SQL types like DATETIME and BIT!!!!
                        throw new NotImplementedException();
                }
            }
            else
            {
                return (null);
            }
        }

        public static SqlDataReader SQLQuery(string qryName, CommandType command)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.LawsonArchive,
                command,
                qryName);
            return (rdr);
        }

        private string SQLGetValueStringFromID(string fieldName, string tableName, string recordIDName, string iDString)
        {
            // Invoke a query of the form "SELECT <fielName> FROM <tableName> WHERE <recordIDname> = <idstring>".
            //   Return the selected record's field value or a blank string if no record is found.
            string strQuery = "";
            try
            {
                // Need to check if the iDString is empty or non-integer, which can/will happen.
                if (iDString.Length == 0)
                {
                    return ("");
                }
                bool idokay = int.TryParse(iDString, out int IDValue);
                if ((!idokay) || (IDValue <= 0))
                {
                    return ("");
                }

                strQuery = "SELECT " + fieldName + " FROM " + tableName + " WHERE " + recordIDName + " = " + iDString;
                using (SqlDataReader rdr = SQLQuery(strQuery, CommandType.Text))
                {
                    // If we obtained a valid record, return the field value from the fieldName parameter
                    if (rdr.HasRows)
                    {
                        rdr.Read();
                        string valuestring = SQLGetString(rdr, fieldName);
                        return (valuestring);
                    }
                    else
                    {
                        return ("");
                    }
                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR in query " + strQuery, ex);
                return ("");
            }
        }

        /// <summary>
        /// A utility to return any value from a SQL Query (as long as the underlying SQL type is knnown apriori).
        /// USAGE:  object x = (T)GetSQLValue(SQLReader, FieldName)
        /// NOTE:   Declarations that use this function MUST be declared nullable!  (i.e., int? var1, double? var2, etc.
        /// </summary>
        /// <param name="rdr"></param>
        /// <param name="s"></param>
        /// <returns>(T)Value</returns>
        public object SQLGetValue(SqlDataReader rdr, string s)
        {
            // A utility to return any value from a SQL Query (as long as the underlying SQL type is knnown apriori).
            // USAGE:
            //   <type T> x = (T)GetSQLValue(SQLReader, FieldName)

            // Because SQL can return a dbnull, there is no way to determine the actual value type.  
            //   This requires that all declarations must be nullable.
            if (rdr[s] != null)
            {
                string t = rdr[s].GetType().ToString().ToLower();
                switch (t)
                {
                    case "system.string":
                        try { return rdr[s].ToString(); } catch { return ""; }
                    case "system.int32":
                        try
                        {
                            string i = rdr[s].ToString();
                            bool iokay = int.TryParse(i, out int ii);
                            return (iokay ? ii : 0);
                        }
                        catch { return 0; }
                    case "system.decimal":
                        try
                        {
                            string d = rdr[s].ToString();
                            bool dokay = Double.TryParse(d, out double dd);
                            return (dokay ? dd : 0);
                        }
                        catch { return 0; }
                    case "system.float":
                        try
                        {
                            string f = rdr[s].ToString();
                            bool fokay = float.TryParse(f, out float ff);
                            return (fokay ? ff : 0);
                        }
                        catch { return 0; }
                    case "system.bit":
                        try
                        {
                            string b = rdr[s].ToString();
                            bool iokay = int.TryParse(b, out int bb);
                            return (iokay ? (bb == 0 ? false : true) : false);
                        }
                        catch { return false; }
                    case "system.bool":
                        try
                        {
                            string b = rdr[s].ToString();
                            bool iokay = int.TryParse(b, out int bb);
                            return (iokay ? (bb == 0 ? false : true) : false);
                        }
                        catch { return false; }
                    case "system.dbnull":
                        // Because SQL can return a dbnull, there is no way to determine the actual value type.  
                        //   This requires that all declarations must be nullable.
                        return null;
                    default:
                        // Check other SQL types like DATETIME and BIT!!!!
                        string e = "Well, lookie here, seems like y'all forgot to handle SQL value type " + t + ". Best be goin' back to the programmer and have 'im do some fixin'";
                        BroadcastError(e, null);
                        return null;
                }
            }
            else
            {
                return (null);
            }
        }

        private string SQLGetString(SqlDataReader rdr, string s)
        {
            try
            {
                return rdr[s].ToString();
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate string from SQL:  " + ex.ToString(), JobName);
                return "";
            }
        }

        private int SQLGetInt(SqlDataReader rdr, string s)
        {
            try
            {
                string a = rdr[s].ToString();
                bool aokay = int.TryParse(a, out int v);
                return (aokay ? v : 0);
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate integer from SQL:  " + ex.ToString(), JobName);
                return 0;
            }
        }

        /// <summary>
        /// Send the specified message (can be of any JobLogMessageType) to the log and to a user prompt.
        /// </summary>
        /// <param name="errorType"></param>
        /// <param name="msg"></param>
        /// <param name="ex"></param>
        private void BroadcastMessage(Enums.JobLogMessageType errorType, string msg, Exception ex)
        {
            // Useful (saves typing) when we want to send the same error message to both the job log and to a user prompt.  Useless otherwise.
            string ExceptionStr = (ex != null ? ex.ToString() : "");
            DataIO.WriteToJobLog(errorType, msg + "\r\n\r\n" + ExceptionStr, JobName);
            MessageBox.Show(msg + ExceptionStr);
        }

        private void BroadcastError(string msg, Exception ex)
        {
            // Useful (saves typing) when we want to send the same error message to both the job log and to a user prompt.  Useless otherwise.
            BroadcastMessage(Enums.JobLogMessageType.ERROR, msg, ex);
        }

        private void BroadcastWarning(string msg, Exception ex)
        {
            // Useful (saves typing) when we want to send the same error message to both the job log and to a user prompt.  Useless otherwise.
            BroadcastMessage(Enums.JobLogMessageType.WARNING, msg, ex);
        }

        private void BroadcastInfo(string msg, Exception ex)
        {
            // Useful (saves typing) when we want to send the same error message to both the job log and to a user prompt.  Useless otherwise.
            BroadcastMessage(Enums.JobLogMessageType.INFO, msg, ex);
        }

        /// <summary>
        /// Clear all text boxes from a panel
        /// </summary>
        /// <param name="p"></param>
        void ClearPanelTextBoxes(Panel p)
        {
            try
            {
                foreach (TextBox t in p.Controls)
                {
                    t.Text = "";
                    t.ForeColor = Color.Black;
                }

            }
            catch { }
        }

        /// <summary>
        /// Clears a combo box without deleting its list items.
        /// </summary>
        /// <param name="cmb"></param>
        private void ClearCombo(object cmb)
        {
            // This clears a combo box without deleting its list items.
            // NOTE that we'll probably throw an exception if we try to clear something other than a combo box
            try
            {
                ((ComboBox)cmb).Text = String.Empty;
                ((ComboBox)cmb).ForeColor = Color.Black;
                ((ComboBox)cmb).SelectedIndex = -1;
                ((ComboBox)cmb).SelectedValue = -1;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to clear combobox " + ((ComboBox)cmb).Name, ex);
            }
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        private void PopulateEmployeeComboBox(bool byEmployeeName)
        {
            // This function constructs the following query:
            //    SELECT * FROM <tablename> ORDER BY <display member> (for all active records)
            // and uses it as the data source for the specified combo box.  
            // displayMember must be a valid field within the dataset, and it is the field that will be displayed in the combo box.

            ComboBox cmb = (ComboBox)CmbEmployee;
            try
            {
                // We probably should convert this to a list and save in a class - we ALSO need to save the ID
                //   as part of each combo box entry (can that be done???  YES!!! It's being done in the "DataComplete" Event handler.
                string SelectSTR = "";
                if (byEmployeeName)
                {
                    SelectSTR = "SELECT CONCAT(RTRIM(LAST_NAME), ', ', RTRIM(FIRST_NAME), ' ', RTRIM(MIDDLE_INIT)) AS NAME, EMPLOYEE FROM [Prod8].[lawson].EMPLOYEE ORDER BY LAST_NAME, FIRST_NAME";
                }
                else
                {
                    SelectSTR = "SELECT CONCAT(EMPLOYEE, ' - ', RTRIM(LAST_NAME), ', ', RTRIM(FIRST_NAME), ' ', RTRIM(MIDDLE_INIT)) AS NAME, EMPLOYEE FROM [Prod8].[lawson].EMPLOYEE ORDER BY EMPLOYEE";

                }
                using (SqlDataReader rdr = SQLQuery(SelectSTR, CommandType.Text))
                {
                    if (rdr.HasRows)
                    {
                        DataTable dt = new DataTable();
                        dt.Load(rdr);
                        cmb.DataSource = dt;
                        cmb.DisplayMember = "NAME";
                        cmb.ValueMember = "EMPLOYEE";
                        cmb.SelectedIndex = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying to populate combobox " + cmb.Name, ex);
            }
        }

        #region CustommenuStripRenderer
        class CustomMenuStripRenderer : ToolStripProfessionalRenderer
        {
            public CustomMenuStripRenderer() : base() { }
            public CustomMenuStripRenderer(ProfessionalColorTable table) : base(table) { }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                e.TextFormat &= ~TextFormatFlags.HidePrefix;
                base.OnRenderItemText(e);
            }
        }
        #endregion

        private void RadEmployeID_CheckedChanged(object sender, EventArgs e)
        {
            PopulateEmployeeComboBox(RadEmployeeName.Checked ? true : false);
        }

        private void RadEmployeeName_CheckedChanged(object sender, EventArgs e)
        {
            PopulateEmployeeComboBox(RadEmployeeName.Checked ? true : false);
        }

        private void CmbEmployee_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool EmployeeIdOkay = false;
            int EmployeeID = 0;
            try
            {
                EmployeeIdOkay = int.TryParse(CmbEmployee.SelectedValue.ToString(), out EmployeeID);
            }
            catch
            {
                // This will happen during initialization.
                return;
            }
            if (EmployeeIdOkay)
            {
                // TBD Run some checks on these values
                // Get the Employee data from the database
                SqlParameter[] EmployeeParams = new SqlParameter[1];
                EmployeeParams[0] = new SqlParameter("@pvintEmployeeNum", EmployeeID);
                List<Dictionary<string, object>> EmployeeData = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.LawsonArchive, CommandType.StoredProcedure, "Proc_Select_Employee_Data", EmployeeParams);

                bool YearOkay = int.TryParse(CmbYear.Text, out int wageyear);
                SqlParameter[] WageParams = new SqlParameter[2];
                WageParams[0] = new SqlParameter("@pvintEmployeeNum", EmployeeID);
                WageParams[1] = new SqlParameter("@pvintPayrollYear", wageyear);
                List<Dictionary<string, object>> WageData = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.LawsonArchive, CommandType.StoredProcedure, "Proc_Select_Wages", WageParams);

                //Display Employee Data
                DisplayEmployeeData(EmployeeData);

                //Display Wage Data
                DisplayWageData(WageData);
            }
        }

        private void CmbYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            CmbEmployee_SelectedIndexChanged(sender, e);
        }

        private void DisplayEmployeeData(List<Dictionary<string, object>> employeeData)
        {
            // Display all employee data in the employee data text box.
            //   The text box intentionally uses fixed-size font (courier) to allow simple text formatting and alignment.
            // Side-by-side columns:  Label (max length = 20) | Value (max length has to fit an entire address, so maybe 80-100)

            Dictionary<string, object> MaritalStatus = new Dictionary<string, object>
            {
                { " ", " " },
                { "S", "Single" },
                { "M", "Married" },
                { "D", "Divorced" }
            };

            Dictionary<string, object> YesNoUnk = new Dictionary<string, object>
            {
                { " ", " " },
                { "Y", "Yes" },
                { "N", "No" },
                { "U", "Unknown" }
            };

            Dictionary<string, object> HourlySalary = new Dictionary<string, object>
            {
                { " ", " " },
                { "H", "Hourly" },
                { "S", "Salaried" }
            };

            Dictionary<string, object> e = employeeData[0];
            EmployeeDataClass edc = new EmployeeDataClass();
            edc.ProduceNameString("Name", e["FIRST_NAME"], e["MIDDLE_INIT"], e["LAST_NAME"]);
            edc.OutputLine("Employee ID", e["EMPLOYEE"].ToString());
            edc.OutputLine("Nickname", e["NICK_NAME"].ToString());
            edc.OutputLine("Gender", e["SEX"].ToString());
            edc.ProduceNameString("Maiden Name", "", "", e["MAIDEN_LST_NM"]);
            edc.ProduceNameString("Former Last Name", "", "", e["FORMER_LST_NM"]);
            edc.ProduceAddress("Address", e["ADDR1"], e["ADDR2"], e["ADDR3"], e["ADDR4"], e["CITY"], e["STATE"], e["ZIP"], e["COUNTRY_CODE"]);
            edc.ProduceAddress("Supplemental Address", e["SUPP_ADDR1"], e["SUPP_ADDR2"], e["SUPP_ADDR3"], e["SUPP_ADDR4"], e["SUPP_CITY"], e["SUPP_STATE"], e["SUPP_ZIP"], e["SUPP_CNTRY_CD"]);
            edc.ProduceComboTwo("Employee Status", e["EMP_STATUS"], e["EMP_STATUS_DESCRIPTION"]);
            edc.ProducePhoneNumber("Home Phone", e["HM_PHONE_CNTRY"], e["HM_PHONE_NBR"], "");
            edc.ProducePhoneNumber("Work Phone", e["WK_PHONE_CNTRY"], e["WK_PHONE_NBR"], e["WK_PHONE_EXT"]);
            edc.ProducePhoneNumber("Suppemental Phone", e["SUPP_PHONE_CNT"], e["SUPP_PHONE_NBR"], "");
            edc.OutputLine("SS #", "TBD");  // Social Security # is TBD - needs to be encrypted
            edc.ProduceComboTwo("Process Level", e["PROCESS_LEVEL"], e["PROCESS_LEVEL_NAME"]);
            edc.ProduceComboTwo("Department", e["DEPARTMENT"], e["DEPARTMENT_NAME"]);
            edc.ProduceComboTwo("Job Code", e["JOB_CODE"], e["JOB_CODE_DESCRIPTION"]);
            edc.OutputLine("Union Code", e["UNION_CODE"].ToString());
            edc.OutputLine("EEO Class", e["EEO_CLASS"].ToString());
            edc.ProduceComboTwo("Marital Status", e["TRUE_MAR_STAT"], (MaritalStatus.FirstOrDefault(k => k.Key == e["TRUE_MAR_STAT"].ToString()).Value));
            edc.ProduceComboTwo("Deceased?", e["DECEASED"], (YesNoUnk.FirstOrDefault(k => k.Key == e["DECEASED"].ToString()).Value));
            edc.ProduceComboTwo("Handicap ID?", e["HANDICAP_ID"], (YesNoUnk.FirstOrDefault(k => k.Key == e["HANDICAP_ID"].ToString()).Value));
            edc.ProduceComboTwo("Veteran?", e["VETERAN"], (YesNoUnk.FirstOrDefault(k => k.Key == e["VETERAN"].ToString()).Value));
            edc.ProduceDate("Birth Date", e["BIRTHDATE"], e["BIRTH_CNTRY_CD"]);
            edc.ProduceDate("Hire Date", e["DATE_HIRED"], "");
            edc.ProduceDate("Adjusted Hire Date", e["ADJ_HIRE_DATE"], "");
            edc.ProduceDate("Anniversary Date", e["ANNIVERS_DATE"], "");
            edc.ProduceDate("Termination Date", e["TERM_DATE"], "");
            edc.ProduceDate("Creation Date", e["CREATION_DATE"], "");
            edc.ProduceDate("Senior Date", e["SENIOR_DATE"], "");
            edc.ProduceDate("Benefit Date", e["BEN_DATE_1"], "");
            edc.ProduceComboTwo("HM Account - Unit", e["HM_ACCOUNT"], e["HM_ACCT_UNIT"]);
            edc.OutputLine("Shift", e["SHIFT"].ToString());
            edc.ProduceComboTwo("Pay Frequency", e["PAY_FREQUENCY"], "Weekly");
            edc.ProduceComboTwo("Salary Class", e["SALARY_CLASS"], (HourlySalary.FirstOrDefault(k => k.Key == e["SALARY_CLASS"].ToString()).Value));
            edc.ProduceComboTwo("Exempt?", e["EXEMPT_EMP"], (YesNoUnk.FirstOrDefault(k => k.Key == e["EXEMPT_EMP"].ToString()).Value));
            edc.ProduceCurrency("Pay Rate", e["PAY_RATE"]);
            edc.ProduceCurrency("Pro Rate Total", e["PRO_RATE_TOTAL"]);
            edc.ProduceCurrency("Pro Rate A Salary", e["PRO_RATE_A_SAL"]);
            edc.ProduceCurrency("Benefit Salary 1", e["BEN_SALARY_1"]);
            edc.ProduceCurrency("Benefit Salary 5", e["BEN_SALARY_5"]);
            edc.ProduceComboTwo("Pension Plan?", e["PENSION_PLAN"], (YesNoUnk.FirstOrDefault(k => k.Key == e["PENSION_PLAN"].ToString()).Value));
            edc.OutputLine("# FTEs", e["NBR_FTE"].ToString());
            edc.OutputLine("FTE Total", e["FTE_TOTAL"].ToString());
            edc.ProduceComboTwo("Auto Time Rec", e["AUTO_TIME_REC"], (YesNoUnk.FirstOrDefault(k => k.Key == e["AUTO_TIME_REC"].ToString()).Value));
            edc.OutputLine("Last Ded Seq", e["LAST_DED_SEQ"].ToString());
            edc.OutputLine("Auto Deposit", e["AUTO_DEPOSIT"].ToString());
            edc.OutputLine("Sec Level", e["SEC_LVL"].ToString());
            edc.OutputLine("Location Code", e["LOCAT_CODE"].ToString());

            // Output the result
            TxtEmployeeData.Text = edc.GetOutputText();
        }

        public class WageDataClass
        {
            const int LABEL_LEN = 5;
            const int VALUE_LEN = 20;

            private string OutputText;

            public WageDataClass()
            {
                OutputText = "";
            }

            internal void ProduceRow(object label, object wages, object currency)
            {
                // Assume that we need room for as much as a hundred thousand dollars
                //  ($100,000.00).  Format so that we have right-alignment on all currency

                const int LABEL_LEN = 5;
                const int HOURS_LEN = 9;
                const int CURRENCY_LEN = 10;
                string s = "|" + (label.ToString()).PadLeft(LABEL_LEN);
                s = s.PadRight(2 * LABEL_LEN);
                string a = "";

                bool hoursokay = double.TryParse(wages.ToString(), out double h);
                if (hoursokay)
                {
                    a = Convert.ToDecimal(h).ToString("#,##0.00");
                    a = " | " + a.PadLeft(HOURS_LEN);
                    s += a;
                }

                bool currencyokay = double.TryParse(currency.ToString(), out double c);
                if (currencyokay)
                {
                    a = Convert.ToDecimal(c).ToString("#,##0.00");
                    a = " | $" + a.PadLeft(CURRENCY_LEN);
                    s += a;
                }


                if (s.Trim().Length > 0)
                {
                    OutputLine("", s);
                }
            }

            public void OutputLine(string label, string val)
            {
                string s = (label.Trim()).PadRight(LABEL_LEN) + (val.Trim().PadRight(VALUE_LEN) + "\r\n");
                OutputText += s;
            }

            internal string GetOutputText()
            {
                return (OutputText);
            }
        }
        public class EmployeeDataClass
        {
            const int LABEL_LEN = 20;
            const int VALUE_LEN = 70;

            private string OutputText;
            private string PreviousAddress;

            public EmployeeDataClass()
            {
                OutputText = "";
                PreviousAddress = ""; // This is a kludge to do an easy comparison between 2 sequential addresses.
            }

            public void ProduceNameString(object label, object first, object middle, object last)
            {
                // Create a name string First Middle Last.  Leave a single blank between the three parts (trim extraneous blanks)
                string s = first.ToString().Trim() + " " + middle.ToString().Trim() + " " + last.ToString().Trim();
                RegexOptions options = RegexOptions.None;
                Regex rgx = new Regex("[ ]{2,}", options);
                s = rgx.Replace(s, " ");
                // Output only if the string is non-blank
                if (s.Trim().Length > 0)
                {
                    OutputLine(label.ToString(), s);
                }
            }

            public void OutputLine(string label, string val)
            {
                string s = (label.Trim()).PadRight(LABEL_LEN) + " | " + (val.Trim().PadRight(VALUE_LEN) + "\r\n");
                OutputText += s;
            }

            public string GetOutputText()
            {
                return (OutputText);
            }

            internal void ProduceAddress(object label, object addr1, object addr2, object addr3, object addr4, object city, object state, object zip, object country)
            {
                string s = "";
                // Produce an address.  Don't bother to use any non-blank objects
                string a = addr1.ToString().Trim();
                if (a.Length > 0) s += a.Trim() + " ";
                a = addr2.ToString().Trim();
                if (a.Length > 0) s += a.Trim() + " ";
                a = addr3.ToString().Trim();
                if (a.Length > 0) s += a.Trim() + " ";
                a = addr4.ToString().Trim();
                if (a.Length > 0) s += a.Trim();
                if (s.Trim().Length > 0) s = s.Trim() + ", ";

                a = city.ToString().Trim();
                if (a.Length > 0) s += a.Trim();
                if (s.Trim().Length > 0) s = s.Trim() + ", ";
                a = state.ToString().Trim();
                if (a.Length > 0) s += a.Trim() + "  ";
                a = zip.ToString().Trim();
                if (a.Length > 0) s += a.Trim() + " ";
                a = country.ToString().Trim();
                if (a.ToUpper() != "US")
                {
                    if (a.Length > 0) s += a.Trim();
                }

                if (s.Trim().Length > 0)
                {
                    // Before saving, compare to the previously-loaded address.  If they are the same then don't reload
                    if (PreviousAddress.ToLower() != s.Trim().ToLower())
                    {
                        OutputLine(label.ToString(), s);
                        PreviousAddress = s.Trim();
                    }
                }

            }

            internal void ProduceComboTwo(string label, object item1, object item2)
            {
                string s = "";
                string a = item1.ToString().Trim();
                if (a.Length > 0) s += a.Trim() + " - ";
                if (!(item2 is null))
                {
                    a = item2.ToString().Trim();
                    if (a.Length > 0) s += a.Trim();
                }

                if (s.Trim().Length > 0)
                {
                    OutputLine(label.ToString(), s);
                }
            }

            internal void ProduceDate(string label, object itemdate, object modifier)
            {
                string s = "";
                string a = itemdate.ToString().Trim();
                bool dateokay = DateTime.TryParse(a, out DateTime d);
                if (dateokay)
                {
                    if (a != "1/1/1753 12:00:00 AM")
                    {
                        s = d.ToString("MM/dd/yyyy");
                    }
                    else
                    {
                        s = "-";
                    }
                    a = modifier.ToString().Trim();
                    if (a.Length > 0)
                    {
                        s += " - " + a;
                    }

                    if (s.Trim().Length > 0)
                    {
                        OutputLine(label.ToString(), s);
                    }
                }
            }

            internal void ProducePhoneNumber(string label, object areacode, object telephonenum, object extension)
            {
                // Concatenate area code and telephone number
                string s = "";
                string a = areacode.ToString().Trim() + telephonenum.ToString().Trim();
                // Remove hyphens if any
                a = a.Replace("-", string.Empty);
                // If Length > 0 (i.e., not empty)...
                if (a.Length > 0)
                {
                    // Re-hyphen after character 3 and if the number is 11 characters, after character 7
                    a = a.Insert(3, "-");
                    if (a.Length == 11) a = a.Insert(7, "-");
                    s = a;

                    // Add extension if any
                    a = extension.ToString().Trim();
                    if (extension.ToString().Length > 0) s += " x" + a;

                    if (s.Trim().Length > 0)
                    {
                        OutputLine(label.ToString(), s);
                    }
                }

            }

            internal void ProduceCurrency(string label, object currency)
            {
                const int CURRENCY_LEN = 10;

                string s = "-";
                bool currencyokay = double.TryParse(currency.ToString(), out double c);
                if (currencyokay)
                {
                    s = Convert.ToDecimal(c).ToString("#,##0.00");
                    s = "$" + s.PadLeft(CURRENCY_LEN);
                }

                if (s.Trim().Length > 0)
                {
                    OutputLine(label.ToString(), s);
                }
            }
        }

        private void DisplayWageData(List<Dictionary<string, object>> wageData)
        {
            // Wages must include summary wage data at the top,
            //    followed by aggregate totals below that

            
            

            WageDataClass wdc = new WageDataClass();

            double TotalHours = 0;
            double TotalWages = 0;

            for (int i = 0; i < wageData.Count; i++)
            {
                bool hoursokay = double.TryParse(wageData[i]["SUMHOURS"].ToString(), out double h);
                bool wagesokay = double.TryParse(wageData[i]["SUMWAGES"].ToString(), out double w);
                if (hoursokay) TotalHours += h;
                if (wagesokay) TotalWages += w;
            }

            wdc.OutputLine("", "");
            wdc.ProduceRow("Yr. Total", TotalHours.ToString(), TotalWages.ToString());
            wdc.OutputLine("", "");

            wdc.OutputLine("", "|Pay Group |   Hours   |    Wages");
            for (int i = 0; i < wageData.Count; i++)
            {
                wdc.ProduceRow(wageData[i]["PAY_SUM_GRP"], wageData[i]["SUMHOURS"], wageData[i]["SUMWAGES"]);
            }

            // Output the result
            TxtWages.Text = wdc.GetOutputText();
        }


        private void MnuPrint_Click(object sender, EventArgs e)
        {
            printdialog.Document = document;
            if (printdialog.ShowDialog() == DialogResult.OK)
            {
                document.Print();
            }
        }

        private void MnuPrintPreview_Click(object sender, EventArgs e)
        {
            previewdialog.Document = document;
            if (previewdialog.ShowDialog() == DialogResult.OK)
            {
                document.Print();
            }
        }

        private void MnuExit_Click(object sender, EventArgs e)
        {
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job completed", JobName);
            Application.Exit();
        }

    }


}
