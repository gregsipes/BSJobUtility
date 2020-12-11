using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BSGlobals;

// Populate an SBS Excel Journal Entry Database with data from exported SBS Excel Reports.
//    This app can read both XLS and XLSX files but the XLS files are limited to 65K rows maximum (as per Microsoft); XLSX export is recommended
//    Convert the app to tab-delimited format so we can use a BULK INSERT to import the data into a "Work" table
//    This app uses a BULK INSERT, so the import MUST EXACTLY CONFORM to the CSV columns
//    Once the insert is complete we need to translate some columns (Money in particular) from text to the appropriate data type
//    Once this is complete we need to copy the data from the "Work" table into the production table (tblJournalEntries),
//       making sure that we remove any duplicates resulting from a re-import of existing data.

namespace SBSJournalEntryImport
{
    public partial class FrmMain : Form
    {

        #region Declarations
        // Constants
        const string JobName = "<Jobname>";

        // Class declarations

        // Other global stuff
        ActiveDirectory UserInfo;
        VersionStatusBar StatusBar;
        string WorkingFolder = "";

        #endregion

        #region Initialization
        public FrmMain()
        {
            InitializeComponent();

            // Job log start
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job starting", JobName);

            // Get configuration values.  

            WorkingFolder = Config.GetConfigurationKeyValue("SBSJournalEntryImport", "Workingfolder");

            // Create event handlers if any 
            //EXAMPLE: TxtAddressLine1.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);

            // Menu strip initialization (where needed)
            //MainMenuStrip.Renderer = new CustomMenuStripRenderer();

            // Get the current (logged-in) username.  It will be in the form DOMAIN\username
            UserInfo = new ActiveDirectory();

            // Add status bar (2 segment default, with version)
            StatusBar = new VersionStatusBar(this);

        }
        #endregion

        #region Data Display Functions

        #endregion

        #region Button Rendering Functions

        #endregion

        #region ----Safe Value Assignments

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

        #endregion

        #region ----SQL

        public static SqlDataReader SQLQuery(string qryName, CommandType command)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.ISInventory,
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

        #endregion

        #region ----General Helper Functions

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

        #region CustommenuStripRenderer
        /// <summary>
        /// 
        /// </summary>
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

        private void CmdBrowse_Click(object sender, EventArgs e)
        {
            // Open up the file open dialog and prompt for an excel file.
            // The file can be anywhere but the CSV-translated result MUST be on a drive that is shared with the SQL server running this app (Omaha)

            try
            {
                OpenFileDialog dlg = new OpenFileDialog
                {
                    Title = "Open Excel Document",
                    //InitialDirectory = @"C:\",
                    Filter = "Documents (Spreadsheets (*.xls,*.xlsx)|*.xls;*.xlsx",
                    RestoreDirectory = true
                };
                DialogResult = dlg.ShowDialog();
                if (DialogResult == DialogResult.OK)
                {
                    TxtImportFile.Text = dlg.FileName;
                    CmdImport.Enabled = true;
                }
                else
                {
                    TxtImportFile.Text = "";
                    CmdImport.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to browse for file", ex);
                return;
            }

            // Convert the selected Excel file to tab-delimited, CSV format.
            //   This will convert any currency values > 999.99 to double-quoted strings, with comma separators in the thousands (and millions) 
            

            Spreadsheet import = new Spreadsheet(TxtImportFile.Text);
            string CSVFilename = WorkingFolder + " SBSJournalEntry_" + UserInfo.Username + ".txt";
            import.File.SaveAs(CSVFilename, Spreadsheet.FileFormat.TabDelimited, true);
            import.Terminate();

            // Truncate, then bulk insert the CSV file into table tblJournalEntries_Imported
            //   IMPORTANT:  The stored procedure assumes that the first N rows of the file are header rows.  This MUST be specified here.

            int NumHeaderRowsInFile = 8;
            SqlParameter[] BulkParams = new SqlParameter[2];
            BulkParams[0] = new SqlParameter("strCSVFilename", CSVFilename);
            BulkParams[1] = new SqlParameter("strNumHeaderRows", NumHeaderRowsInFile.ToString());
            DataIO.ExecuteNonQuery(Enums.DatabaseConnectionStringNames.SBSJournalEntryImport, CommandType.StoredProcedure, "Proc_BulkInsertSBSJournalEntries", BulkParams);

            // Transfer the data from the imported table to the target table (tblJournalEntries),
            //   performing any data translations along the way.  
            //   As of 12/10/2020 the only known translation is Amount (text) to currency.
            // We also have to overwrite any existing identical records lest we create duplicates




        }

        private void CmdExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
