using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using BSGlobals;

// A maintenance app that allows easy access to re-trying a failed SYNERGY report.
//   After selecting the report this app will update the appropriate Loads record
//   in the PBS2Macro database so that the (automated) PBS2Macro SQL Job will
//   rerun the report (which will occur within, a most, a 5-minute interval).

namespace Synergy_Report_Maintenance
{    
    public partial class FrmMain : Form
    {

        const string JobName = "SyergyReportMaintenance";
        ActiveDirectory UserInfo = new ActiveDirectory();
        VersionStatusBar StatusBar;

        public FrmMain()
        {
            InitializeComponent();
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job starting", JobName);

            ADUserClass ADList = UserInfo.ADUserList[0];
            if (!((ADList.HasCredential("synergy_accounting")) || (ADList.HasCredential("admin_synergy")) || (ADList.HasCredential("information systems"))))
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR - User needs Synergy_Accounting, Admin_Synergy or Information Systems credentials to use this app", JobName);
                MessageBox.Show("ERROR - You need Synergy_Accounting, Admin_Synergy or Information Systems credentials to use this app");
                Application.Exit();
            }

            // Initialize controls
            CmdRefresh.Enabled = false;

            // Add status bar (2 segment default, with version)
            StatusBar = new VersionStatusBar(this);
        }

        /// <summary>
        /// Populates the selected combobox.
        /// </summary>
        /// <param name="cmb">Name of the combobox</param>
        /// <param name="tableName">Name of the table whose data will be loaded into the combobox</param>
        /// <param name="displayMember">Name of the field in the table to be displayed</param>
        /// <param name="valueMember">Name of the primary key field (must be a unique ID)</param>
        private void PopulateComboBox(ComboBox cmb, string SelectStr, string displayMember, string valueMember)
        {
            // This function makes a query out of the passed-in string:
            // and uses it as the data source for the specified combo box.  
            // displayMember must be a valid field within the dataset, and it is the field that will be displayed in the combo box.

            try
            {
                using (SqlDataReader rdr = SQLQuery(SelectStr, CommandType.Text))
                {
                    if (rdr.HasRows)
                    {
                        DataTable dt = new DataTable();
                        dt.Load(rdr);
                        cmb.DataSource = dt;
                        cmb.DisplayMember = displayMember;
                        cmb.ValueMember = valueMember;
                        cmb.SelectedIndex = -1;
                    }
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate combobox " + cmb.Name + ": " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate combobox " + cmb.Name + ": " + ex.ToString());
            }
        }

        #region SQL

        public static SqlDataReader SQLQuery(string qryName, CommandType command)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.SynergyReportMaintenance,
                command,
                qryName);
            return (rdr);
        }

        public static SqlDataReader SQLQuery(string qryName)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.SynergyReportMaintenance,
                CommandType.StoredProcedure,
                qryName);
            return (rdr);
        }

        public static SqlDataReader SQLQuery(string qryName, SqlParameter[] orderParams)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.SynergyReportMaintenance,
                CommandType.StoredProcedure,
                qryName,
                orderParams);
            return (rdr);
        }

        public static void SQLProcCall(string procName, SqlParameter[] Params)
        {
            DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.SynergyReportMaintenance,
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



        #endregion



        private void CmbSynergyReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            // This is a fixed-entry dropdown list in the following order (by index)
            // 0 - AGEDET
            // 1 - AGESUMM
            // 2 - AGESUMM2
            // 3 - DLYDRAW
            // 4 - DEPOSITS
            // 5 - DEPOSIT
            // 6 - PREAUTH
            // 7 - GRACEOWE
            // 8 - GRACEWO
            // 9 - UR

            // Load the CMBReportToRefresh dropdown with the latest (last couple of months) entries from the Loads table
            int index = CmbSynergyReportType.SelectedIndex;
            string WhereStr = "";
            DateTime ThreeMonthsAgo = DateTime.Now.AddMonths(-3);
            switch (index)
            {
                case 0: 
                    WhereStr = "AND original_file LIKE '%AGEDET.%' ";
                    break;
                case 1:
                    WhereStr = "AND original_file LIKE '%AGESUMM.%' ";
                    break;
                case 2:
                    WhereStr = "AND original_file LIKE '%AGESUMM2.%' ";
                    break;
                case 3:
                    WhereStr = "AND original_file LIKE '%DLYDRAW.%' ";
                    break;
                case 4:
                    WhereStr = "AND original_file LIKE '%DEPOSITS.%' ";
                    break;
                case 5:
                    WhereStr = "AND original_file LIKE '%DEPOSIT.%' ";
                    break;
                case 6:
                    WhereStr = "AND original_file LIKE '%PREAUTH.%' ";
                    break;
                case 7:
                    WhereStr = "AND original_file LIKE '%GRACEOWE.%' ";
                    break;
                case 8:
                    WhereStr = "AND original_file LIKE '%GRACEWO.%' ";
                    break;
                case 9:
                    WhereStr = "AND original_file LIKE '%UR.%' ";
                    break;
            }
            string SelectStr = "SELECT Original_file FROM Loads WHERE Load_Date > '" + ThreeMonthsAgo.ToShortDateString() + "' " + WhereStr + " ORDER BY load_date DESC ";
            PopulateComboBox(CmbReportToRefresh, SelectStr, "Original_File", "Original_File");
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job Completed", JobName);
        }

        private void CmbReportToRefresh_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Once the appropriate file is selected, enable the RefreshReport button
            CmdRefresh.Enabled = true;
        }

        private void CmdRefresh_Click(object sender, EventArgs e)
        {
            // Take the selected file and rename it with today's date/time so it no longer appears in its original form
            //   in this table.  SQL job PBS2Macro will re-process the original file if it still exists in \\circfs\spoolcm\normal.
            DateTime now = DateTime.Now;
            string datetime = now.Year.ToString("D4") + now.Month.ToString("D2") + now.Day.ToString("D2") + "_" + now.Hour.ToString("D2") + now.Minute.ToString("D2") + now.Second.ToString("D2");
            string newfilename = CmbReportToRefresh.Text + "_" + datetime;
            string UpdateStr = "UPDATE Loads SET Original_File = '" + newfilename + "' WHERE Original_File = '" + CmbReportToRefresh.Text + "' ";
            SQLQuery(UpdateStr, CommandType.Text);
        }

        private void CmdExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
