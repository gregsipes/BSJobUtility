using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using BSGlobals;

namespace IS_Inventory
{
    public partial class FrmMain : Form
    {
        #region General

        #region ----General Declarations
        const string JobName = "IS Inventory";
        ActiveDirectory UserInfo = new ActiveDirectory();
        VersionStatusBar StatusBar;
        bool IsInitializing;

        #endregion

        #region ----General Initialization

        public FrmMain()
        {
            InitializeComponent();

            IsInitializing = true;
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job starting", JobName);

            InitializeHardwareTemplate();
            InitializeCategoryEdit();
            InitializeIPAddressesEdit();
            InitializePCsAndMACs();

            // Add status bar (2 segment default, with version)
            StatusBar = new VersionStatusBar(this);

            // The TabInventory_Layout event will occur immediately after we leave this function.
        }

        private void TabInventory_Layout(object sender, LayoutEventArgs e)
        {
            // This is only needed to initialize the first (visible) tab.
            //  However, this event is invoked a large number of times during loading/initialization.  Luckily this process is very slim.
            if (IsInitializing)
            {
                TabInventory.SelectedIndex = 0;
                TabInventory_SelectedIndexChanged(sender, e);
            }

            // This is the last event during initialization.  Mark initialization as complete.
            IsInitializing = false;
        }
        #endregion

        #region ----General Classes

        /// <summary>
        /// A general status class that provides simple status related to a SQL-populated dictionary
        /// </summary>
        public class DictionaryStatusClass
        {
            public List<Dictionary<string, object>> FieldList; // Conforms to DataIO dictionary return type.
            public bool IsSaved;
            public bool IsNew;
            public bool IsActive;
            public int SelectedIndex;

            public DictionaryStatusClass()
            {
                FieldList = new List<Dictionary<string, object>>();
                IsSaved = true;
                IsNew = true;
                IsActive = true;
                SelectedIndex = -1;
            }
        }
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

#region ----SQL

        public static SqlDataReader SQLQuery(string qryName, CommandType command)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.ISInventory,
                command,
                qryName);
            return (rdr);
        }

        public static SqlDataReader SQLQuery(string qryName)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.ISInventory,
                CommandType.StoredProcedure,
                qryName);
            return (rdr);
        }

        public static SqlDataReader SQLQuery(string qryName, SqlParameter[] orderParams)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.ISInventory,
                CommandType.StoredProcedure,
                qryName,
                orderParams);
            return (rdr);
        }

        public static void SQLProcCall(string procName, SqlParameter[] Params)
        {
            DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory,
                CommandType.StoredProcedure,
                procName,
                Params);
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
            MessageBox.Show(msg + "\r\n\r\n" + ExceptionStr);
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
                for (int i = 0; i < p.Controls.Count; i++)
                {
                    if (p.Controls[i] is TextBox t)
                    {
                        t.Text = "";
                        t.ForeColor = Color.Black;
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Sets the .IsSaved property in the associated status object to false (indicating that the class is dirty)
        ///    and enables the associated SAVE button
        /// </summary>
        /// <param name="status"></param>
        /// <param name="button"></param>
        private void MarkAsDirty(object sender, DictionaryStatusClass status, Button button)
        {
            // Generic DateChanged event whenever a record gets dirty.
            //  Useful (so far) for text boxes, combo boxes and date/time picker boxes.

            if (sender.GetType() == typeof(TextBox))
            {
                ((TextBox)sender).ForeColor = Color.Red;
            }

            if ((sender.GetType() == typeof(ComboBox)) || (sender.GetType() == typeof(ComboBoxUnlocked)))            {
                ((ComboBox)sender).ForeColor = Color.Red;
            }
            if (sender.GetType() == typeof(DateTimePicker))
            {
                // DateTimePicker control's forecolor or backcolor attributes do not work.  Use a label instead
                ((DateTimePicker)sender).ForeColor = Color.Red;
            }

            if (!(status == null))
            {
                SetButtonState(button, true, Color.White, Color.Red);
                status.IsSaved = false;
            }
        }

        private void SetButtonState(Button button, Boolean enabled, Color forecolor, Color backcolor)
        {
            button.Enabled = enabled;
            button.BackColor = backcolor;
            button.ForeColor = forecolor;
        }

        /// <summary>
        /// Executed when the user hits the ENTER key or when focus moves off the control.  
        /// </summary>
        /// <param name="cmb"></param>
        /// <returns></returns>
        private bool AddServerComboboxEntry(ComboBox cmb)
        {
            // User Control ComboBoxUnlocked ONLY
            // 
            // Invoked when the user hits the ENTER key OR when focus moves off the ComboBoxUnlocked control.
            //   When this happens we want to check if the user entered a text value not already on the combobox list.
            //   Anything not already on the list deserves a prompt asking if it should be added to the list.
            //   Once it's added we need to reload the combo box.
            // NOTE:  The Combobox items list is not IEnumerable so LINQ doesn't work directly on this list.
            //   While there are tricky ways around this it's just as easy to query the underyling table to see
            //   if the item is already in the table.

            // Don't process if no text was entered
            if (cmb.Text.Length == 0) return (false);

            bool recordadded = false;
            string newdata = cmb.Text;
            try
            {
                switch (cmb.Name.ToLower())
                {
                    case "cmbpccabinets_id": recordadded = AddToTableIfNotAMember(cmb, "lstCabinets", "Cabinet", "Cabinets_ID", newdata); break;
                    case "cmbpccddevices_id": recordadded = AddToTableIfNotAMember(cmb, "lstCDDevices", "CD_Device", "CD_Devices_ID", newdata); break;
                    case "cmbservercontainertype": recordadded = AddToTableIfNotAMember(cmb, "lstContainerTypes", "container_type", "container_types_id", newdata); break;
                    case "cmbpcdepartment": recordadded = AddToTableIfNotAMember(cmb, "lstDepartments", "department", "departments_id", newdata); break;
                    case "cmbpcharddrive1_id": recordadded = AddToTableIfNotAMember(cmb, "lstDrives", "Drives", "Drives_ID", newdata); break;
                    case "cmbpcharddrive2_id": recordadded = AddToTableIfNotAMember(cmb, "lstDrives", "Drives", "Drives_ID", newdata); break;
                    case "cmbpcmiscdrives_id": recordadded = AddToTableIfNotAMember(cmb, "lstDrives", "Drives", "Drives_ID", newdata); break;
                    case "cmbpckeyboards_id": recordadded = AddToTableIfNotAMember(cmb, "lstKeyboards", "Keyboard", "Keyboards_ID", newdata); break;
                    case "cmbpcmanufacturers_id": recordadded = AddToTableIfNotAMember(cmb, "lstManufacturers", "Manufacturer", "Manufacturers_ID", newdata); break;
                    case "cmbservermanufacturer": recordadded = AddToTableIfNotAMember(cmb, "lstManufacturers", "Manufacturer", "Manufacturers_ID", newdata); break;
                    case "cmbpcmice_id": recordadded = AddToTableIfNotAMember(cmb, "lstMice", "Mouse", "Mice_ID", newdata); break;
                    case "cmbpcmiscellaneouscard_id": recordadded = AddToTableIfNotAMember(cmb, "lstMiscellaneous", "Miscellaneous", "Miscellaneous_ID", newdata); break;
                    case "cmbpcmodels_id": recordadded = AddToTableIfNotAMember(cmb, "lstModels", "Model", "Models_ID", newdata); break;
                    case "cmbservermodel": recordadded = AddToTableIfNotAMember(cmb, "lstModels", "Model", "Models_ID", newdata); break;
                    case "cmbpcmonitor1_id": recordadded = AddToTableIfNotAMember(cmb, "lstMonitors", "Monitor", "Monitors_ID", newdata); break;
                    case "cmbpcmonitor2_id": recordadded = AddToTableIfNotAMember(cmb, "lstMonitors", "Monitor", "Monitors_ID", newdata); break;
                    case "cmbpcmotherboards_id": recordadded = AddToTableIfNotAMember(cmb, "lstMotherboards", "Motherboard", "Motherboards_ID", newdata); break;
                    case "cmbpcnics_id": recordadded = AddToTableIfNotAMember(cmb, "lstNICs", "NIC", "NICs_ID", newdata); break;
                    case "cmbserverowner": recordadded = AddToTableIfNotAMember(cmb, "lstOwners", "owner_name", "owners_id", newdata); break;
                    case "cmbpctype": recordadded = AddToTableIfNotAMember(cmb, "lstPCTypes", "pctype", "PCType_id", newdata); break;
                    case "cmbpcprocessors_id": recordadded = AddToTableIfNotAMember(cmb, "lstProcessors", "Processor", "Processors_ID", newdata); break;
                    case "cmbserverprocessor": recordadded = AddToTableIfNotAMember(cmb, "lstProcessors", "Processor", "Processors_ID", newdata); break;
                    case "cmbpcram_id": recordadded = AddToTableIfNotAMember(cmb, "lstRAM", "RAM", "RAM_ID", newdata); break;
                    case "cmbserverram": recordadded = AddToTableIfNotAMember(cmb, "lstRAM", "RAM", "RAM_ID", newdata); break;
                    case "cmbpcsoundcards_id": recordadded = AddToTableIfNotAMember(cmb, "lstSoundCards", "Sound_Card", "Sound_Cards_ID", newdata); break;
                    case "cmbpcspeakers_id": recordadded = AddToTableIfNotAMember(cmb, "lstSpeakers", "Speakers", "Speakers_ID", newdata); break;
                    case "cmbpcvideocards_id": recordadded = AddToTableIfNotAMember(cmb, "lstVideoCards", "Video_Card", "Video_Cards_ID", newdata); break;
                    default:
                        BroadcastWarning("ERROR trying to process combobox " + cmb.Name + " with data " + newdata + " (combobox not listed in switch statement)", null);
                        break;
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR processing combobox " + cmb.Name + ", data = " + newdata, ex);
            }
            return (recordadded);
        }

        private bool AddToTableIfNotAMember(ComboBox cmb, string tableName, string fieldName, string recIDName, string fieldValue)
        {
            // Assume that the specified entry is NOT in the table, AND assume that the user doesn't not want to add it (i.e., it's a typo).
            bool returnstatus = false;
            try
            {
                string SelectStr = "SELECT * FROM " + tableName + " WHERE " + fieldName + " = '" + fieldValue + "'";
                using (SqlDataReader rdr = SQLQuery(SelectStr, CommandType.Text))
                {
                    if (!rdr.HasRows)
                    {
                        DialogResult dlgresult = MessageBox.Show("'" + fieldValue + "' is not on this list.  Add it?", "NOT ON LIST", MessageBoxButtons.YesNo);
                        if (dlgresult == DialogResult.Yes)
                        {
                            string InsertStr = "INSERT INTO " + tableName + "(" + fieldName + ") VALUES ('" + fieldValue + "')";
                            SQLQuery(InsertStr, CommandType.Text);
                            PopulateComboBox(cmb, tableName, fieldName, recIDName);
                            cmb.Text = fieldValue;
                            returnstatus = true;
                        }
                    }
                    else
                    {
                        // This record is already in the selected table so just return true (it's already a member)
                        returnstatus = true;
                    }
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR adding value to list table " + tableName + ", field " + fieldName, ex);
            }
            return (returnstatus);
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

        /// <summary>
        /// Populates the selected combobox.  Throws an exception if anything other than a combobox invokes this function.
        /// </summary>
        /// <param name="cmb">Name of the combobox</param>
        /// <param name="tableName">Name of the table whose data will be loaded into the combobox</param>
        /// <param name="displayMember">Name of the field in the table to be displayed</param>
        /// <param name="valueMember">Name of the primary key field (must be a unique ID)</param>
        private void PopulateComboBox(object combobox, string tableName, string displayMember, string valueMember)
        {
            // This function constructs the following query:
            //    SELECT * FROM <tablename> ORDER BY <display member> (for all active records)
            // and uses it as the data source for the specified combo box.  
            // displayMember must be a valid field within the dataset, and it is the field that will be displayed in the combo box.

            ComboBox cmb = (ComboBox)combobox;
            try
            {
                // We probably should convert this to a list and save in a class - we ALSO need to save the ID
                //   as part of each combo box entry (can that be done???  YES!!! It's being done in the "DataComplete" Event handler.
                string SelectSTR = "SELECT * FROM " + tableName + " WHERE active_flag = 1 ORDER BY " + displayMember;
                using (SqlDataReader rdr = SQLQuery(SelectSTR, CommandType.Text))
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
                BroadcastError("ERROR trying to populate combobox " + cmb.Name, ex);
            }
        }

        /// <summary>
        /// Displays query results on a generic grid based on active/inactive/both display status
        /// (If clearDataSource is true, lets Windows decide column sizes and row heights, 
        ///   and displays all fields in the query).
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="query"></param>
        void PopulateGrid(DataGridView dgv, string query, CommandType command, RadioButton activeRadioButton)
        {
            try
            {
                SqlDataReader rdr;
                SqlParameter[] Params = null; // This is the default for text-based queries
                string QueryStr = query;

                // Check if a radio button prefix argument was specified.  If so, then we need to include a WHERE clause
                //   as part of the query - selecting either Active or Inactive subsets of the list with which to populate the grid.
                // NOTE: For this to be usable, ALL groups of radio buttons must have the form 
                //   <buttonname>Active     <buttonname>Inactive     <buttonname>both
                // What gets passed in is the ACTIVE button name; we'll strip it down to its prefix in order to form the "inactive" button name.

                if (activeRadioButton != null)
                {
                    string buttonprefix = activeRadioButton.Name.Left(activeRadioButton.Name.Length - 6); // strip off "Active"
                    RadioButton inactiveRadioButton = (RadioButton)activeRadioButton.Parent.Controls[buttonprefix + "Inactive"];

                    switch (command)
                    {
                        case CommandType.StoredProcedure:
                            // If the commandtype is a stored procedure, then use a SqlParameter to specify argument "@pvintIsActive"
                            //    NOTE: The stored procedure MUST have an argument names "@pvintIsActive"!!!!!
                            QueryStr = query;
                            Params = new SqlParameter[1];
                            if (activeRadioButton.Checked || inactiveRadioButton.Checked)
                            {
                                Params[0] = new SqlParameter("@pvintIsActive", activeRadioButton.Checked ? 1 : 0);
                            }
                            else
                            {
                                // Neither Active nor Inactive radio button was checked, so assume that there's a "Both" button and that is was checked.
                                //   Therefore, no argument needs to be passed into the stored procedure (the null default will be taken)
                                Params[0] = new SqlParameter("@pvintIsActive", null);  // TBD Need to confirm that this works
                            }
                            break;
                        case CommandType.Text:
                            // If a query was passed in, append the active/inactive state as part of a where clause.
                            //   This is a bit tougher because we need to discover if "WHERE" was already inserted into the string,
                            //   at which point we'll replace it with "WHERE (active_flag = xxxx) AND " to maintain where clause validity.

                            if (activeRadioButton.Checked || inactiveRadioButton.Checked)
                            {
                                string SelectStr = query.ToLower();
                                int index = SelectStr.IndexOf(" where ");
                                if (index > 0)
                                {
                                    string leftside = SelectStr.Left(index);
                                    string rightside = SelectStr.Right(SelectStr.Length - (index + 7));
                                    string clause = " WHERE (active_flag = " + (activeRadioButton.Checked ? "1" : "0") + ") AND ";
                                    QueryStr = leftside + clause + rightside;
                                }
                                else
                                {
                                    // No WHERE clause is included in the query, so append it to the query
                                    QueryStr = query + " WHERE active_flag = " + (activeRadioButton.Checked ? "1" : "0");
                                }
                            }
                            else
                            {
                                // Neither Active nor Inactive radio button was checked, so assume that there's a "Both" button and that is was checked.
                                //   Therefore, no argument needs to be passed into the stored procedure (the null default will be taken)
                                QueryStr = query;
                            }
                            break;
                        case CommandType.TableDirect:
                            // Not used
                            QueryStr = query;  // This is just for completeness in the hopes of not throwing an exception.
                            break;
                    }
                }
                else
                {
                    // If no radio buttons are specified then we have to rely purely on the passed-in query (or stored procedure)
                    //  without any additional filtering
                }

                // Populate the specified grid using a generic approach - unilaterally fill it left to right
                //   based on the order specified in the stored procedure.

                using (rdr = DataIO.ExecuteQuery(Enums.DatabaseConnectionStringNames.ISInventory, command, QueryStr, Params))
                {
                    //if (rdr.HasRows)
                    {
                        dgv.Visible = true;
                        DataTable dt = new DataTable();
                        dt.Load(rdr);
                        dgv.DataSource = null; // using blanks ("") wipes out any named columns in the grid
                        dgv.DataSource = dt;
                    }
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to populate grid " + dgv.Name + " with " + query, ex);
            }
        }

#endregion

#region ----General Events
        private void FrmMain_Load(object sender, EventArgs e)
        {
            // TBD - Nothing yet needed here so it's just a placeholder for now
        }

        private void TabInventory_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                switch (TabInventory.SelectedIndex)
                {
                    case 0:  // PCs and MACs
                             // On select, populate the PC / Mac tab
                        if (!PCsAndMacsInitialized)
                        {
                            PCsAndMacsInitialized = true;
                            PopulatePCsAndMacsGrid();
                            PopulateComboBox(CmbPCType, "lstPCTypes", "pctype", "PCType_id");
                            PopulateComboBox(CmbPCDepartment, "lstDepartments", "department", "departments_id");
                            PopulateComboBox(CmbPCTemplate, "TemplatesHardware", "template_name", "hardware_templates_id"); // WARNING: This is not the same design as the lst tables but it works for this table
                        }
                        // 
                        break;

                    case 1:  // IP Addresses
                             // On select, populate the IP Address tab
                        if (!IPAddressesInitialized)
                        {
                            IPAddressesInitialized = true;
                            PopulateIPAddressGrid();
                        }
                        //GrdIPAddresses.Sort(GrdIPAddresses.Columns[GridSortColumn], GridSortOrder);   // Sort this by IPAddress (at least, initially and after every insert)
                        break;
                    case 2:  // Servers
                        if (!ServersInitialized)
                        {
                            ServersInitialized = true;
                            ServerRecord = new ServerEditClass();  
                            PopulateServersGrid();
                            ServerRecord.Status.IsSaved = true;
                        }
                        break;
                    case 3:  // Documentation
                        if (!DocumentationInitialized)
                        {
                            DocumentationInitialized = true;
                            DocumentationRecord = new DocumentationClass();
                            PopulateDocumentationGrid();
                            SetButtonState(CmdDocumentationSearch, false, Color.Black, Color.White);
                        }
                        break;
                    case 4: // HW Template Edit
                            // Nothing to do here, move along
                        break;
                    case 5: // Category Edit
                            // Nothing to do here, move along
                        break;
                    default:
                        BroadcastError("ERROR: Unhandled tab index" + TabInventory.SelectedIndex.ToString() + " in TabInventory_SelectedIndexChanged event", null);
                        break;

                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR: Tab processing error in TabInventory_SelectedIndexChanged", ex);
            }

        }
        private void SubtabTemplateEdit_Layout(object sender, LayoutEventArgs e)
        {
            // On layout, populalate the selected editor's comboboxes

            string tabname = e.AffectedControl.Name.ToLower();
            switch (tabname)
            {
                case "subtabhardware":
                    PopulateHardwareTemplateTab();
                    break;
            }
        }
#endregion

#endregion

#region Hardware Template Tab

#region ----HW Tab Declarations
        HardwareTemplateClass HWTemplate;
        bool HardwareTemplateTabInitialized = false;

        private void InitializeHardwareTemplate()
        {
            PnlHardwareTemplate.Visible = false;
        }

#endregion

#region ----Populate HW Tab Templates
        /// <summary>
        /// Populates the Hardware Template Tab.
        /// </summary>
        private void PopulateHardwareTemplateTab()
        {
            // Hardware Template
            // Do this only when this tab gets the focus
            try
            {
                // ONLY DO THIS ONCE PER SESSION.
                if (HardwareTemplateTabInitialized) return;

                // Combobox event handlers.
                // This will allow us to point the same event handler to ALL hardware tab ComboBoxUnlocked controls.
                //   We may want separate event handlers for various tabs/sections, or at least the ability to only
                //   create the handlers when the user opens the tab.
                foreach (Control c in SubtabTemplateEdit.TabPages["SubtabHardware"].Controls["PnlHardwareTemplate"].Controls)
                {
                    if (c is ComboBoxUnlocked)
                    {
                        ((ComboBoxUnlocked)c).DataEntryComplete += new EventHandler(Event_DataEntryComplete);
                    }
                }

                PopulateComboBox(CmbPCmanufacturers_id, "lstManufacturers", "Manufacturer", "Manufacturers_ID");
                PopulateComboBox(CmbPCmodels_id, "lstModels", "Model", "Models_ID");
                PopulateComboBox(CmbPCcabinets_id, "lstCabinets", "Cabinet", "Cabinets_ID");
                PopulateComboBox(CmbPCprocessors_id, "lstProcessors", "Processor", "Processors_ID");
                PopulateComboBox(CmbPCmotherboards_id, "lstMotherboards", "Motherboard", "Motherboards_ID");
                PopulateComboBox(CmbPCmonitor1_id, "lstMonitors", "Monitor", "Monitors_ID");
                PopulateComboBox(CmbPCmonitor2_id, "lstMonitors", "Monitor", "Monitors_ID");
                PopulateComboBox(CmbPCmice_id, "lstMice", "Mouse", "Mice_ID");
                PopulateComboBox(CmbPCnics_id, "lstNICs", "NIC", "NICs_ID");
                PopulateComboBox(CmbPCvideocards_id, "lstVideoCards", "Video_Card", "Video_Cards_ID");
                PopulateComboBox(CmbPCsoundcards_id, "lstSoundCards", "Sound_Card", "Sound_Cards_ID");
                PopulateComboBox(CmbPCcddevices_id, "lstCDDevices", "CD_Device", "CD_Devices_ID");
                PopulateComboBox(CmbPCkeyboards_id, "lstKeyboards", "Keyboard", "Keyboards_ID");
                PopulateComboBox(CmbPCharddrive1_id, "lstDrives", "Drives", "Drives_ID");
                PopulateComboBox(CmbPCharddrive2_id, "lstDrives", "Drives", "Drives_ID");
                PopulateComboBox(CmbPCmiscdrives_id, "lstDrives", "Drives", "Drives_ID");
                PopulateComboBox(CmbPCram_id, "lstRAM", "RAM", "RAM_ID");
                PopulateComboBox(CmbPCspeakers_id, "lstSpeakers", "Speakers", "Speakers_ID");
                PopulateComboBox(CmbPCmiscellaneouscard_id, "lstMiscellaneous", "Miscellaneous", "Miscellaneous_ID");

                PopulateGrid(GrdHardwareTemplate, "Proc_Select_Hardware_Template", CommandType.StoredProcedure, RadHWTemplateFilterActive);
                HardwareTemplateTabInitialized = true;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to populate Hardware Template tab", ex);
            }
            return;
        }
#endregion

#region ----HW Tab Classes

        public class HardwareTemplateClass                                                                         // TBD use the generic DictionaryStatusClass here!!!!!!!!!
        {
            public DictionaryStatusClass Status;

            // A simple class containing a dictionary containing all fields in the TempatesHardware table.
#if false
            public List<Dictionary<string, object>> HWRecord; // Conforms to DataIO dictionary return type.
            public bool IsActive;
            public bool IsSaved;
#endif

            public HardwareTemplateClass()
            {
                Status = new DictionaryStatusClass();
#if false
                HWRecord = new List<Dictionary<string, object>>();
                IsSaved = true;
                IsActive = true;
#endif
                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["hardware_templates_ID"] = 0,
                    ["template_name"] = ""
                };
                Status.FieldList.Add(fields);
            }
        }

#endregion

#region ----HW Tab Data Display Functions
        /// <summary>
        /// Clears the EDIT panel of the Hardware Template tab and populates it with the selected hardware template record.
        /// </summary>
        /// <param name="hWTemplate"></param>
        /// <param name="recordID"></param>
        private void DisplayHardwareTemplate(HardwareTemplateClass hWTemplate, int recordID)
        {
            try
            {
                ClearHWTemplate();
                SqlParameter[] TemplateParams = new SqlParameter[1];
                TemplateParams[0] = new SqlParameter("@pvintTemplateID", recordID);
                hWTemplate.Status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Hardware_Template", TemplateParams);
                Dictionary<string, object> hwr = HWTemplate.Status.FieldList[0];
                SafeText(TxtPCtemplate_name, hwr, "template_name");
                SafeText(CmbPCcabinets_id, hwr, "Cabinet");
                SafeText(CmbPCcddevices_id, hwr, "CD_Device");
                SafeText(CmbPCharddrive1_id, hwr, "Drive1");
                SafeText(CmbPCharddrive2_id, hwr, "Drive2");
                SafeText(CmbPCkeyboards_id, hwr, "Keyboard");
                SafeText(CmbPCmanufacturers_id, hwr, "Manufacturer");
                SafeText(CmbPCmiscdrives_id, hwr, "MiscDrive");
                SafeText(CmbPCmiscellaneouscard_id, hwr, "Miscellaneous");
                SafeText(CmbPCmodels_id, hwr, "Model");
                SafeText(CmbPCmonitor1_id, hwr, "Monitor1");
                SafeText(CmbPCmonitor2_id, hwr, "Monitor2");
                SafeText(CmbPCmotherboards_id, hwr, "Motherboard");
                SafeText(CmbPCmice_id, hwr, "Mouse");
                SafeText(CmbPCnics_id, hwr, "NIC");
                SafeText(CmbPCprocessors_id, hwr, "Processor");
                SafeText(CmbPCram_id, hwr, "RAM");
                SafeText(CmbPCsoundcards_id, hwr, "Sound_Card");
                SafeText(CmbPCspeakers_id, hwr, "Speakers");
                SafeText(CmbPCvideocards_id, hwr, "Video_Card");
                SafeRadioBox(RadPCActive, RadPCInactive, hwr, "active_flag");
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to populate hardware template editor", ex);
            }
        }
#endregion

#region ----HW Tab Helper Functions

        private void SaveHardwareTemplate(HardwareTemplateClass hWTemplate)
        {
            // I thought about redoing this using the HWTemplate's dictionary.  Put everything in a loop.
            //  Unfortunately, there's much more in the dictionary than just the fields in the table,
            //  so that might be a really tough thing to work out.
            try
            {
                SqlParameter[] HWParams = new SqlParameter[23];
                HWParams[0] = new SqlParameter("@pvintTemplateID", hWTemplate.Status.FieldList[0]["hardware_templates_id"]);
                HWParams[1] = new SqlParameter("@pvchrTemplateName", TxtPCtemplate_name.Text);
                HWParams[2] = new SqlParameter("@pvintManufacturersID", CmbPCmanufacturers_id.SelectedValue);
                HWParams[3] = new SqlParameter("@pvintModelsID", CmbPCmodels_id.SelectedValue);
                HWParams[4] = new SqlParameter("@pvintCabinetsID", CmbPCcabinets_id.SelectedValue);
                HWParams[5] = new SqlParameter("@pvintMotherboardsID", CmbPCmotherboards_id.SelectedValue);
                HWParams[6] = new SqlParameter("@pvintProcessorsID", CmbPCprocessors_id.SelectedValue);
                HWParams[7] = new SqlParameter("@pvintRAMID", CmbPCram_id.SelectedValue);
                HWParams[8] = new SqlParameter("@pvintHarddrive1ID", CmbPCharddrive1_id.SelectedValue);
                HWParams[9] = new SqlParameter("@pvintHarddrive2ID", CmbPCharddrive2_id.SelectedValue);
                HWParams[10] = new SqlParameter("@pvintMiscdrivesID", CmbPCmiscdrives_id.SelectedValue);
                HWParams[11] = new SqlParameter("@pvintMonitor1ID", CmbPCmonitor1_id.SelectedValue);
                HWParams[12] = new SqlParameter("@pvintMonitor2ID", CmbPCmonitor2_id.SelectedValue);
                HWParams[13] = new SqlParameter("@pvintVideoCardsID", CmbPCvideocards_id.SelectedValue);
                HWParams[14] = new SqlParameter("@pvintSoundCardsID", CmbPCsoundcards_id.SelectedValue);
                HWParams[15] = new SqlParameter("@pvintSpeakersID", CmbPCspeakers_id.SelectedValue);
                HWParams[16] = new SqlParameter("@pvintNICsID", CmbPCnics_id.SelectedValue);
                HWParams[17] = new SqlParameter("@pvintCDDevicesID", CmbPCcddevices_id.SelectedValue);
                HWParams[18] = new SqlParameter("@pvintMiscellaneousCardID", CmbPCmiscellaneouscard_id.SelectedValue);
                HWParams[19] = new SqlParameter("@pvintKeyboardsID", CmbPCkeyboards_id.SelectedValue);
                HWParams[20] = new SqlParameter("@pvintMiceID", CmbPCmice_id.SelectedValue);
                HWParams[21] = new SqlParameter("@pvchrComments", TxtPCcomments.Text);
                HWParams[22] = new SqlParameter("@pcbitIsActive", RadPCactive_flag.Checked ? 1 : 0);

                HWTemplate.Status.IsSaved = true;
                SQLQuery("Proc_Update_Hardware_Template", HWParams);

                // Refresh the grid
                ClearHWTemplate();
                PopulateGrid(GrdHardwareTemplate, "Proc_Select_Hardware_Template", CommandType.StoredProcedure, RadHWTemplateFilterActive);
                PnlHardwareTemplate.Visible = false;
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully save hardware template", ex);
            }
        }

        private void ClearHWTemplate()
        {
            // Remove all selected information from each of the combo boxes in the HW Template edit panel.
            //   (Start with a clean slate)
            TxtPCtemplate_name.Text = "";
            TxtPCcomments.Text = "";
            TxtPCtemplate_name.ForeColor = Color.Black;
            TxtPCcomments.ForeColor = Color.Black;
            // TBD NEED TO SET TEXT BOXES RED ON Change (COMBOBOXES TOO!)

            ClearCombo(CmbPCmanufacturers_id);
            ClearCombo(CmbPCmodels_id);
            ClearCombo(CmbPCcabinets_id);
            ClearCombo(CmbPCmotherboards_id);
            ClearCombo(CmbPCprocessors_id);
            ClearCombo(CmbPCram_id);
            ClearCombo(CmbPCharddrive1_id);
            ClearCombo(CmbPCharddrive2_id);
            ClearCombo(CmbPCmiscdrives_id);
            ClearCombo(CmbPCmonitor1_id);
            ClearCombo(CmbPCmonitor2_id);
            ClearCombo(CmbPCvideocards_id);
            ClearCombo(CmbPCsoundcards_id);
            ClearCombo(CmbPCspeakers_id);
            ClearCombo(CmbPCnics_id);
            ClearCombo(CmbPCcddevices_id);
            ClearCombo(CmbPCmiscellaneouscard_id);
            ClearCombo(CmbPCkeyboards_id);
            ClearCombo(CmbPCmice_id);
            RadPCactive_flag.Checked = true;
        }

#endregion

#region ----HW Tab Events

        private void CmdNewHardwareTemplate_Click(object sender, EventArgs e)
        {
            // Check if there is a current record and if so, whether or not is has been saved.
            //   If not saved, the prompt to save it.

            try
            {
                if (HWTemplate != null)
                {
                    if (!HWTemplate.Status.IsSaved)
                    {
                        DialogResult result = MessageBox.Show("The current template has not been saved.  Do you wish to save it now?", "RECORD NOT YET SAVED", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            SaveHardwareTemplate(HWTemplate);
                        }
                    }
                }

                // Create a new template.  Name it "HW Template #nnn" where nnn is the ID #
                ClearHWTemplate();

                string templatename = "HW Template #";
                SqlParameter[] HWParams = new SqlParameter[1];
                HWParams[0] = new SqlParameter("@pvchrTemplateName", templatename);
                SqlDataReader rdr = SQLQuery("Proc_Insert_Hardware_Template", HWParams);

                // Construct a (temporary) name for this template and display it
                rdr.Read();
                int RecID = SQLGetInt(rdr, "hardware_templates_id");
                TxtPCtemplate_name.Text = templatename + RecID.ToString();
                HWTemplate = new HardwareTemplateClass();
                HWTemplate.Status.FieldList[0]["hardware_templates_id"] = RecID;
                HWTemplate.Status.FieldList[0]["template_name"] = TxtPCtemplate_name.Text;

                // Update the new record with the (temporary) template name
                HWParams = new SqlParameter[2];
                HWParams[0] = new SqlParameter("@pvintTemplateID", RecID);
                HWParams[1] = new SqlParameter("@pvchrTemplateName", TxtPCtemplate_name.Text);
                SQLQuery("Proc_Update_Hardware_Template", HWParams);

                PnlHardwareTemplate.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully select hardware template", ex);
            }
        }

        /// <summary>
        ///  Close the log for this entry
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job Completed", JobName);
        }

        /// <summary>
        /// Executed whenever a grid row is clicked.  Display that row on the HW Template's edit panel.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdHardwareTemplate_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;

            // If a new template entry has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            try
            {
                if (HWTemplate != null)
                {
                    if (!HWTemplate.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE TEMPLATE?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes:
                                // Save record first  
                                if (TxtPCtemplate_name.Text.Length > 0)
                                {
                                    SaveHardwareTemplate(HWTemplate);
                                }
                                else
                                {
                                    MessageBox.Show("Sorry, a template name must be entered before saving.  Do tis first (or next time, answer NO to the previous SAVE prompt)");
                                    return;
                                }
                                break;
                            case DialogResult.No:
                                // Fall through this switch and load the data onto the edit window.
                                break;
                            case DialogResult.Cancel:
                                // Do NOTHING regarding the row click
                                return;
                        }
                    }
                }

                // Load the selected grid data into the edit template.  We need the IDs of all of the fields to make this work!
                HWTemplate = new HardwareTemplateClass();
                PnlHardwareTemplate.Visible = true;
                DisplayHardwareTemplate(HWTemplate, (int)GrdHardwareTemplate.Rows[row].Cells["hardware_templates_ID"].Value);
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfull process hardware template grid selection", ex);
            }
        }


        /// <summary>
        /// An event unique to ComboBoxUnlocked.  It fires when the user depresses the ENTER key or when the combobox loses focus.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Event_DataEntryComplete(object sender, EventArgs e)
        {
            try
            {
                ComboBox cmb = (ComboBox)sender;

                // If this entry isn't on the list then add it
                bool recordadded = AddServerComboboxEntry(cmb);

                // If the combobox entry was accepted (added to the table or already a member of the table) then update its dictionary entry in the template record
                // (fieldname is identical to combobox name EXCEPT for the leading 5 letters (in this case, "CMBPC" needs to be stripped)
                if (recordadded)
                {
                    string keyname = cmb.Name.Right(cmb.Name.Length - 5);
                    HWTemplate.Status.FieldList[0][keyname] = cmb.SelectedValue;
                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR: Unable to successful process combobox entry", ex);
            }
        }

        /// <summary>
        /// Save the current HW template edit panel to the TemplatesHardware table
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmbSaveHardwareTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                // The only item in a hardware template that we need is the template's name.  It must be unique
                if (TxtPCtemplate_name.Text.Length > 0)
                {
                    SaveHardwareTemplate(HWTemplate);
                }
                else
                {
                    MessageBox.Show("Please enter a template name before saving");
                    return;
                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully save hardware template", ex);
            }
        }

        private void RadPCactive_flag_CheckedChanged(object sender, EventArgs e)
        {
            // Mark whether or not the record is rendered active.  Inactive records won't appear in other lists.
            HWTemplate.Status.IsActive = RadPCactive_flag.Checked;
        }

        private void RadHWTemplateActiveStatus_Changed(object sender, EventArgs e)
        {
            PopulateGrid(GrdHardwareTemplate, "Proc_Select_Hardware_Template", CommandType.StoredProcedure, RadHWTemplateFilterActive);
        }

#endregion

#endregion

#region Category Edit

#region ----Category Edit Declarations
        CategoryEditClass CategoryEdit;

        private void InitializeCategoryEdit()
        {
            PnlCategoryEditNewSave.Visible = false;
            PnlCategoryEditItem.Visible = false;
            SetButtonState(CmdCategoryEditSaveItem, false, Color.Black, Color.White);
        }
#endregion

#region ----Category Edit Classes

        public class CategoryEditClass
        {
            // A simple class containing all fields in the lst<xxx> categories tables.
            public bool IsSaved;
            public bool IsNew;
            public int RecordID;
            public string ItemName;
            public bool IsActive;
            public string LastModifiedBy;

            public string TableName;

            public string FieldRecordIDName;
            public string FieldItemName;
            public string FieldIsActiveName;
            public string FieldLastModifiedByName;

            public CategoryEditClass(DataGridView dgv, string tableName)
            {
                IsSaved = true;
                IsNew = true;
                RecordID = 0;
                ItemName = "";
                IsActive = true;
                LastModifiedBy = "";

                TableName = "lst" + tableName;

                FieldRecordIDName = dgv.Columns[0].Name;
                FieldItemName = dgv.Columns[1].Name;
                FieldIsActiveName = dgv.Columns[2].Name;
                FieldLastModifiedByName = dgv.Columns[3].Name;
            }
        }
#endregion

#region ----Category Edit Data Display Functions
        /// <summary>
        /// Clears the EDIT panel of the Hardware Template tab and populates it with the selected hardware template record.
        /// </summary>
        /// <param name="hWTemplate"></param>
        /// <param name="recordID"></param>
        private void DisplayCategoryEdit(CategoryEditClass categoryEdit, int recordID)
        {
            try
            {
                TxtItemName.Text = categoryEdit.ItemName;
                if (categoryEdit.IsActive)
                {
                    RadCategoryActive.Checked = true;
                }
                else
                {
                    RadCategoryInactive.Checked = true;
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to populate Category editor", ex);
            }
        }

        private void ClearCategoryEditTemplate()
        {
            // Remove all selected information from the Category edit panel.
            //   (Start with a clean slate)
            TxtItemName.Text = "";
            TxtItemName.ForeColor = Color.Black;
            RadCategoryActive.Checked = true;
        }

#endregion

#region ----Category Edit Events
        private void RadCategoryActiveStatus_Changed(object sender, EventArgs e)
        {
            // Used by the Show Active / Show Inactive / Show Both radio buttons, this event forces a save to the selected record.
            if (CmbSelectedList.Text.Length > 0)
            {
                CmbSelectedList_SelectedIndexChanged(sender, e);
            }
        }

        private void CmbSelectedList_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Populate the grid with whatever selection was made here
            SelectCategoryListToPopulate();

            //Render the New/Save panel visible
            PnlCategoryEditNewSave.Visible = true;

            // TBD THERE IS NO STATUS CLASS FOR CATEGORIES, so we need to come up with a different method for marking selections red/black
        }

        private void SelectCategoryListToPopulate()
        {
            try
            {
                string listname = CmbSelectedList.Text; // The names in this combobox list MUST be exactly the same as the tablenames, less the "lst" prefix.
                string SelectStr = "SELECT * FROM lst" + listname + " ";
                PopulateGrid(GrdSelectedList, SelectStr, CommandType.Text, RadCategoryListActive);
                FormatCategoryGrid(GrdSelectedList);
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully populate category " + CmbSelectedList.Text, ex);
            }
        }

        private void FormatCategoryGrid(DataGridView dgv)
        {
            // The 4 columns of this grid are always the same:
            //    Record ID (hide it)
            //    Item Name (text)    - read only
            //    IsActive (boolean)  - read only
            //    Last Modified by (hide it)
            dgv.Columns[0].Visible = false;
            dgv.Columns[1].ReadOnly = true;
            dgv.Columns[2].ReadOnly = true;
            dgv.Columns[3].Visible = false;
        }

        private void CmdCategoryEditNewItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Check if a previous record needs to be saved before creating a new record!
                CheckForEditedCategoryItem(CategoryEdit);

                // Create a new category record and initialize the editor
                CategoryEdit = new CategoryEditClass(GrdSelectedList, CmbSelectedList.Text);
                FormatCategoryGrid(GrdSelectedList);
                ClearCategoryEditTemplate();
                PnlCategoryEditItem.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully create new Category item", ex);
            }
        }

        private void GrdSelectedList_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;

            try
            {
                // If a new item entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.
                CheckForEditedCategoryItem(CategoryEdit);

                // Populate the Category Edit panel with the selected item contents
                CategoryEdit = new CategoryEditClass(GrdSelectedList, CmbSelectedList.Text);
                FormatCategoryGrid(GrdSelectedList);
                DataGridViewRow dgr = GrdSelectedList.Rows[row];
                CategoryEdit.RecordID = (int)dgr.Cells[0].Value;
                CategoryEdit.ItemName = dgr.Cells[1].Value.ToString();
                CategoryEdit.IsActive = (bool)dgr.Cells[2].Value;
                CategoryEdit.LastModifiedBy = dgr.Cells[3].Value.ToString();
                CategoryEdit.IsNew = false;

                DisplayCategoryEdit(CategoryEdit, (int)GrdSelectedList.Rows[row].Cells[0].Value);
                PnlCategoryEditItem.Visible = true;
                SetButtonState(CmdCategoryEditSaveItem, false, Color.Black, Color.White);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully process grid selection", ex);
            }
        }

        private void CheckForEditedCategoryItem(CategoryEditClass categoryEdit)
        {
            if (CategoryEdit != null)
            {
                if (!CategoryEdit.IsSaved)
                {
                    DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE TEMPLATE?", MessageBoxButtons.YesNoCancel);
                    switch (dlgResult)
                    {
                        case DialogResult.Yes: SaveCategoryEdit(CategoryEdit); break;
                        case DialogResult.No: break;
                        case DialogResult.Cancel: return;
                    }
                }
            }
        }

        private void CmdCategoryEditSaveItem_Click(object sender, EventArgs e)
        {
            // The only item in the editor that we need is the item name.
            if (TxtItemName.Text.Length > 0)
            {
                SaveCategoryEdit(CategoryEdit);
            }
            else
            {
                MessageBox.Show("Please enter an item name before saving");
                return;
            }

        }

        private void SaveCategoryEdit(CategoryEditClass categoryEdit)
        {
            try
            {
                string UpdateStr = "";
                if (categoryEdit.IsNew)
                {
                    UpdateStr =
                        "INSERT INTO " + categoryEdit.TableName +
                        "(" + categoryEdit.FieldItemName + ", " + categoryEdit.FieldIsActiveName + ", " + categoryEdit.FieldLastModifiedByName + ") " +
                        "VALUES ('" + TxtItemName.Text + "', " + (RadCategoryActive.Checked ? "1" : "0") + ", '" + categoryEdit.LastModifiedBy + "')";
                }
                else
                {
                    UpdateStr =
                        "UPDATE " + categoryEdit.TableName + " SET " +
                        categoryEdit.FieldItemName + " = '" + TxtItemName.Text + "', " +
                        categoryEdit.FieldIsActiveName + " = " + (RadCategoryActive.Checked ? "1" : "0") + ", " +
                        categoryEdit.FieldLastModifiedByName + " = '" + categoryEdit.LastModifiedBy + "' " +
                        "WHERE " + categoryEdit.FieldRecordIDName + " = " + categoryEdit.RecordID;
                }

                SQLQuery(UpdateStr, CommandType.Text);

                // Refresh the grid
                ClearCategoryEditTemplate();
                SelectCategoryListToPopulate();
                PnlCategoryEditItem.Visible = false;
                SetButtonState(CmdCategoryEditSaveItem, false, Color.Black, Color.White);
                categoryEdit.IsSaved = true;
                categoryEdit.IsNew = false;
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to save category selection to " + categoryEdit.TableName, ex);
            }
        }

        private void TxtItemName_TextChanged(object sender, EventArgs e)
        {
            // This text box is limited to 100 characters - same length as maximum allowed in the category list tables
            SetButtonState(CmdCategoryEditSaveItem, true, Color.White, Color.Red);
            CategoryEdit.IsSaved = false;  // TBD NEED TO CHANGE THIS TO A STATUS OBJECT!!!!!
            // TBD We want to use MarkAsDirty here!!!!!
        }

        private void RadCategoryActive_CheckedChanged(object sender, EventArgs e)
        {
            SetButtonState(CmdCategoryEditSaveItem, true, Color.White, Color.Red);
            CategoryEdit.IsSaved = false;
        }

#endregion

#endregion

#region IP Address Edit

#region ----IP Address Edit Declarations

        IPAddressEditClass IPAddressRecord;
        ListSortDirection IPAddressGridSortOrder;
        int IPAddressGridSortColumn;
        bool IPAddressesInitialized;

        private void InitializeIPAddressesEdit()
        {
            // By default the SAVE button and the IP address text box are not enabled.
            //   (Existing IP addresses are not editable, and SAVE is enabled only after an edit)

            SetButtonState(CmdSaveIPAddress, false, Color.Black, Color.White);
            TxtIPAddress.Enabled = false;
            PnlIPEditItem.Visible = false;
            IPAddressGridSortOrder = ListSortDirection.Ascending;
            IPAddressGridSortColumn = 1; // Initially sorted by IP Address
            IPAddressesInitialized = false;
        }

#endregion

#region ----IP Address Edit Classes

        public class IPAddressEditClass
        {
            // A simple class containing all fields in the IPAddresses table
            public DictionaryStatusClass Status;

            public IPAddressEditClass()
            {
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["ipaddresses_id"] = 0,
                    ["IPAddress"] = "",
                    ["VLAN"] = "",
                    ["Description"] = "",
                    ["TranslatedAddr"] = "",
                    ["Notes"] = "",
                    ["IsActive"] = true,
                    ["LastModified"] = "",
                    ["ModifiedBy"] = ""
                };
                Status.FieldList.Add(fields);
            }
        }
#endregion

#region ----IP Address Edit Data Display Functions

        private void PopulateIPAddressGrid()
        {
            try
            {
                PopulateGrid(GrdIPAddresses, "Proc_Select_IP_Addresses", CommandType.StoredProcedure, RadIPAddressesFilterActive);
                SetButtonState(CmdSaveIPAddress, false, Color.Black, Color.White);

                // For this grid we need to specify column attributes
                GrdIPAddresses.Columns[0].Visible = false;   // Hide the IP Address ID
                GrdIPAddresses.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;  // The NOTES column will be fixed length
                GrdIPAddresses.Columns[5].Width = 500;
                GrdIPAddresses.Sort(GrdIPAddresses.Columns[IPAddressGridSortColumn], IPAddressGridSortOrder);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully populate IP Address grid", ex);
            }
        }

        private void DisplayIPAddressRecord(IPAddressEditClass iPAddressRecord, int recordID)
        {
            // Get all data associated with this record
            try
            {
                ClearIPAddressEditPanel();
                SqlParameter[] IPAddressParams = new SqlParameter[1];
                IPAddressParams[0] = new SqlParameter("@pvintRecordID", recordID);
                iPAddressRecord.Status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_IP_Addresses", IPAddressParams);
                Dictionary<string, object> ipr = iPAddressRecord.Status.FieldList[0];

                SafeText(TxtIPAddress, ipr, "IPAddress");
                SafeText(TxtIPVLAN, ipr, "VLAN");
                SafeText(TxtIPDescription, ipr, "Description");
                SafeText(TxtIPTranslatedAddress, ipr, "TranslatedAddr");
                SafeText(TxtIPNotes, ipr, "Notes");
                SafeRadioBox(RadIPActive, RadIPInactive, ipr, "active_flag");

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdSaveIPAddress, false, Color.Black, Color.White);
                iPAddressRecord.Status.IsSaved = true;

                // Existing IP addresses are NOT editable, so disable the IPAddress text box
                TxtIPAddress.Enabled = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to populate IP address record editor", ex);
            }
        }

        private void ClearIPAddressEditPanel()
        {
            // Remove all selected information from each of the text boxes in the IP Address edit panel.
            //   (Start with a clean slate)
            ClearPanelTextBoxes(PnlIPEditItem);
            RadIPActive.Checked = true;
        }

#endregion

#region ----IP Address Edit Helper Functions

        private void SaveIPAddressRecord(IPAddressEditClass iP)
        {
            // Qualification tests:  
            //   IPAddress and IPTranslatedAddress MUST be a valid IPV4 address
            //   VLAN must be a positive integer
            try
            {
                if (QualifyIPAddressValues())
                {

                    // Is this a new record (i.e., is the IP Address text box enabled?) If so, then perform a record insert prior to the update.
                    bool IsNew = false;
                    if (TxtIPAddress.Enabled)
                    {

                        if (!IsDuplicateIPAddress(TxtIPAddress.Text))
                        {
                            try
                            {
                                SqlParameter[] InsertParams = new SqlParameter[4];
                                InsertParams[0] = new SqlParameter("@pvintOctet1", TranslateToOctet(TxtIPAddress.Text, 1));
                                InsertParams[1] = new SqlParameter("@pvintOctet2", TranslateToOctet(TxtIPAddress.Text, 2));
                                InsertParams[2] = new SqlParameter("@pvintOctet3", TranslateToOctet(TxtIPAddress.Text, 3));
                                InsertParams[3] = new SqlParameter("@pvintOctet4", TranslateToOctet(TxtIPAddress.Text, 4));
                                SqlDataReader rdr = SQLQuery("Proc_Insert_IP_Addresses", InsertParams);
                                using (rdr)
                                {
                                    if (rdr.HasRows)
                                    {
                                        rdr.Read();
                                        iP.Status.FieldList[0]["ipaddresses_id"] = SQLGetInt(rdr, "ipaddresses_id");
                                        IsNew = true;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                BroadcastWarning("ERROR trying to create new IP Address record", ex);
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("IP Address " + TxtIPAddress.Text + " is already in the database.  You cannot save a duplicate address.", "DUPLICATE IP ADDRESS", MessageBoxButtons.OK);
                            return;
                        }
                    }

                    SqlParameter[] IPParams = new SqlParameter[15];
                    IPParams[0] = new SqlParameter("@pvintRecordID", iP.Status.FieldList[0]["ipaddresses_id"]);
                    IPParams[1] = new SqlParameter("@pvintOctet1", TranslateToOctet(TxtIPAddress.Text, 1));
                    IPParams[2] = new SqlParameter("@pvintOctet2", TranslateToOctet(TxtIPAddress.Text, 2));
                    IPParams[3] = new SqlParameter("@pvintOctet3", TranslateToOctet(TxtIPAddress.Text, 3));
                    IPParams[4] = new SqlParameter("@pvintOctet4", TranslateToOctet(TxtIPAddress.Text, 4));
                    IPParams[5] = new SqlParameter("@pvchrTranslatedOctet1", TranslateToOctet(TxtIPTranslatedAddress.Text, 1));
                    IPParams[6] = new SqlParameter("@pvchrTranslatedOctet2", TranslateToOctet(TxtIPTranslatedAddress.Text, 2));
                    IPParams[7] = new SqlParameter("@pvchrTranslatedOctet3", TranslateToOctet(TxtIPTranslatedAddress.Text, 3));
                    IPParams[8] = new SqlParameter("@pvchrTranslatedOctet4", TranslateToOctet(TxtIPTranslatedAddress.Text, 4));
                    IPParams[9] = new SqlParameter("@pvchrVLAN", TxtIPVLAN.Text);
                    IPParams[10] = new SqlParameter("@pvchrDescription", TxtIPDescription.Text);
                    IPParams[11] = new SqlParameter("@pvchrNotes", TxtIPNotes.Text);
                    IPParams[12] = new SqlParameter("@pvbitIsActive", RadIPActive.Checked ? 1 : 0);
                    IPParams[13] = new SqlParameter("@pvdatLastModifiedDate", DateTime.Now);
                    IPParams[14] = new SqlParameter("@pvchrLastModifiedBy", UserInfo.Username);

                    SQLQuery("Proc_Update_IP_Addresses", IPParams);

                    // Refresh the grid
                    ClearIPAddressEditPanel();
                    PopulateIPAddressGrid();
                    // Re-sort the grid ONLY if an insert occurred (otherwise, just leave the grid as is)
                    if (IsNew)
                    {
                        //GrdIPAddresses.Sort(GrdIPAddresses.Columns[GridSortColumn], GridSortOrder);   // Sort this by IPAddress (at least, initially and after every insert)
                    }
                    PnlIPEditItem.Visible = false;
                    iP.Status.IsSaved = true;
                }

            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to save new IP Address record", ex);
            }

        }

        private bool IsDuplicateIPAddress(string iPAddress)
        {
            try
            {
                // Check if the passed-in IP Address is already in the IPAddresses table.  Return true if it is.
                string SelectSTR = "SELECT * FROM IPAddresses WHERE " +
                    " octet1 = " + TranslateToOctet(iPAddress, 1) + " AND " +
                    " octet2 = " + TranslateToOctet(iPAddress, 2) + " AND " +
                    " octet3 = " + TranslateToOctet(iPAddress, 3) + " AND " +
                    " octet4 = " + TranslateToOctet(iPAddress, 4);
                using (SqlDataReader rdr = SQLQuery(SelectSTR, CommandType.Text))
                {
                    return (rdr.HasRows);
                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  IsDuplicateIPAddress parsing failure", ex);
                return (false);
            }
        }

        private bool QualifyIPAddressValues()
        {
            // Qualify the following values:
            //  IP address as IPV4
            System.Net.IPAddress addr;
            if (TxtIPAddress.Text.Trim().Length > 0)
            {
                try
                {
                    if (TxtIPAddress.Text.Count(d => d == '.') != 3)
                    {
                        MessageBox.Show("IP Address MUST be a valid IPV4 address (n.n.n.n); please re-enter");
                        return (false);
                    }
                    addr = System.Net.IPAddress.Parse(TxtIPAddress.Text);
                }
                catch
                {
                    MessageBox.Show("IP Address MUST be a valid IPV4 address (n.n.n.n); please re-enter");
                    return (false);
                }
            }

            //  IP translated address as IPV4
            if (TxtIPTranslatedAddress.Text.Trim().Length > 0)
            {
                try
                {
                    if (TxtIPTranslatedAddress.Text.Count(d => d == '.') != 3)
                    {
                        MessageBox.Show("IP Translated Address MUST be a valid IPV4 address (n.n.n.n); please re-enter");
                        return (false);
                    }
                    addr = System.Net.IPAddress.Parse(TxtIPTranslatedAddress.Text);
                }
                catch
                {
                    MessageBox.Show("IP Translated Address MUST be a valid IPV4 address (n.n.n.n); please re-enter");
                    return (false);
                }
            }

            //  VLAN as a positive integer between 0 and 9999
            if (TxtIPVLAN.Text.Trim().Length > 0)
            {
                int x = -1;
                if (!(int.TryParse(TxtIPVLAN.Text, out x) && x > 0 && x < 10000))
                {
                    MessageBox.Show("VLAN value MUST be a positive integer between 1 and 9999; please re-enter");
                    return (false);
                }
            }

            // Everything is hunky-dorey; return true.
            return (true);
        }

        private string TranslateToOctet(string iPAddress, int index)
        {
            // Translate a standard IPV4 address to individual bytes, and return the selected byte.
            // NOTE that the expected indexes range from 1 to 4 and NOT 0 to 3!

            try
            {
                if (iPAddress.Trim().Length > 0)
                {
                    System.Net.IPAddress addr = System.Net.IPAddress.Parse(iPAddress);
                    byte[] b = addr.GetAddressBytes();
                    return (b[index - 1].ToString());
                }
                else
                {
                    return ("");
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR translating IP Address to octets", ex);
                return ("");
            }
        }

#endregion

#region ----IP Address Events

        private void GrdIPAddresses_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction
                IPAddressGridSortColumn = e.ColumnIndex;
                IPAddressGridSortOrder = (GrdIPAddresses.Columns[IPAddressGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            try
            {
                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (IPAddressRecord != null)
                {
                    if (!IPAddressRecord.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE IP ADDRESS RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveIPAddressRecord(IPAddressRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }
                // Load the selected grid data into the edit template. 
                IPAddressRecord = new IPAddressEditClass();
                PnlIPEditItem.Visible = true;
                DisplayIPAddressRecord(IPAddressRecord, (int)GrdIPAddresses.Rows[row].Cells["ipaddresses_id"].Value);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully process IP Address grid selection", ex);
            }
        }

        private void CmdSaveIPAddress_Click(object sender, EventArgs e)
        {
            SaveIPAddressRecord(IPAddressRecord);
        }

        private void IPAddress_TextChanged(object sender, EventArgs e)
        {
            // This event is associated with ALL text and radio boxes in the edit menu.
            //   See the PROPERTIES View for each of the test and radio boxes
            // On any change, enable the Save button
            MarkAsDirty(sender, IPAddressRecord.Status, CmdSaveIPAddress);
        }

        private void CmdNewIPAddress_Click(object sender, EventArgs e)
        {
            try
            {
                // Enable the IPAddress text box
                TxtIPAddress.Enabled = true;

                // On New, render the IP Edit panel
                PnlIPEditItem.Visible = true;

                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (IPAddressRecord != null)
                {
                    if (!IPAddressRecord.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE IP ADDRESS RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveIPAddressRecord(IPAddressRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Create a new record
                IPAddressRecord = new IPAddressEditClass();
                ClearIPAddressEditPanel();
                // BECAUSE the IP Address on a new record MUST be a valid IP Address,
                //   defer creating a new table record until AFTER the user hits the
                //   save button.  Use data entry in the (enabled) IPAddress text box
                //   to determine that we can add this IP address as a new record
                TxtIPAddress.Enabled = true;
                PnlIPEditItem.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully create a new IP Address record", ex);
            }
        }

        private void RadIPAddressesActiveStatus_Changed(object sender, EventArgs e)
        {
            PopulateGrid(GrdIPAddresses, "Proc_Select_IP_Addresses", CommandType.StoredProcedure, RadIPAddressesFilterActive);
        }

#endregion

#endregion

#region PCs and MACs

#region ----PCs and MACs Declarations

        PCsAndMACsEditClass PCsAndMacsRecord;
        ListSortDirection PCsAndMacsGridSortOrder;
        int PCsAndMacsGridSortColumn;
        bool PCsAndMacsInitialized;

        private void InitializePCsAndMACs()
        {
            PnlPCUser.Visible = false;   // Render this after a NEW or an existing selection is made
            PnlPCHardware.Visible = false; // Render this only after a HW template has been selected
            SetButtonState(CmdSavePC, false, Color.Black, Color.White);

            // Populate Department, PC Type and PC Template comboboxes
            PopulateComboBox(CmbPCTemplate, "TemplatesHardware", "template_name", "hardware_templates_id");
            PopulateComboBox(CmbPCType, "lstPCTypes", "pctype", "PCType_id");
            PopulateComboBox(CmbPCDepartment, "lstDepartments", "department", "departments_id");

        }

#endregion

#region ----PCs and MACs Classes
        public class PCsAndMACsEditClass
        {
            // A simple class containing all fields in the IPAddresses table
            public DictionaryStatusClass Status;

            public PCsAndMACsEditClass()
            {
                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["pcmac_id"] = 0,
                    ["system_id"] = "",
                    ["active_flag"] = true
                };
                Status = new DictionaryStatusClass();
                Status.FieldList.Add(fields);
            }
        }

#endregion

#region ----PCs and MACs Data Display Functions

        private void DisplayPCsAndMACsRecord(PCsAndMACsEditClass pCMac, int recordID)
        {
            // Get all data associated with this record
            try
            {
                ClearPCsAndMACsEditPanel();

                SqlParameter[] PCMacParams = new SqlParameter[1];
                PCMacParams[0] = new SqlParameter("@pvintPCMacID", recordID);
                pCMac.Status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_PCsAndMACs", PCMacParams);
                Dictionary<string, object> pcm = pCMac.Status.FieldList[0];

                SafeText(CmbPCType, pcm, "pc_type");
                SafeText(TxtPCSystemID, pcm, "system_id");
                SafeText(TxtPCUsername, pcm, "username");
                SafeText(CmbPCDepartment, pcm, "user_department");
                SafeComboBox(CmbPCTemplate, SQLGetValueStringFromID("template_name", "TemplatesHardware", "hardware_templates_id", pcm["pc_template_id"].ToString()));
                SafeText(TxtPCSystemTag, pcm, "system_tag");
                SafeText(TxtPCSystemSN, pcm, "system_serial_number");
                SafeText(TxtPCMonitorTag, pcm, "monitor_tag");
                SafeText(TxtPCMonitorSN, pcm, "monitor_serial_number");
                SafeText(TxtPCKeyboardTag, pcm, "keyboard_tag");
                SafeText(TxtPCKeyboardSN, pcm, "keyboard_serial_number");
                SafeText(TxtPCMACAddress, pcm, "macaddress");
                SafeText(TxtPCManufacturer, pcm, "manufacturer");
                SafeText(TxtPCCabinet, pcm, "cabinet");
                SafeText(TxtPCModel, pcm, "model");
                SafeText(TxtPCProcessor, pcm, "processor");
                SafeText(TxtPCRAM, pcm, "ram");
                SafeText(TxtPCMotherboard, pcm, "motherboards");
                SafeText(TxtPCMonitor1, pcm, "monitor1");
                SafeText(TxtPCMonitor2, pcm, "monitor2");
                SafeText(TxtPCHardDrive1, pcm, "hard_drive1");
                SafeText(TxtPCHardDrive2, pcm, "hard_drive2");
                SafeText(TxtPCMiscDrive, pcm, "miscellaneous_drive");
                SafeText(TxtPCNIC, pcm, "nic");
                SafeText(TxtPCCDDevice, pcm, "cd_device");
                SafeText(TxtPCVideoCard, pcm, "video_card");
                SafeText(TxtPCSoundCard, pcm, "sound_card");
                SafeText(TxtPCMiscCard, pcm, "miscellaneous_card");
                SafeText(TxtPCSpeakers, pcm, "speakers");
                SafeText(TxtPCKeyboard, pcm, "keyboard");
                SafeText(TxtPCMouse, pcm, "mouse");
                SafeText(TxtPCComments2, pcm, "comments");
                SafeRadioBox(RadPCActive, RadPCInactive, pcm, "active_flag");

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdSavePC, false, Color.Black, Color.White);
                pCMac.Status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR displaying PC / Mac records", ex);
            }
        }

        private void PopulatePCsAndMacsGrid()
        {
            try
            {
                PopulateGrid(GrdPCsAndMacs, "Proc_Select_PCsAndMACs", CommandType.StoredProcedure, RadPCsAndMacsFilterActive);
                SetButtonState(CmdSavePC, false, Color.Black, Color.White);
                // For this grid we need to specify column attributes
                GrdPCsAndMacs.Columns[0].Visible = false;   // Hide the IP Address ID
                GrdPCsAndMacs.Sort(GrdPCsAndMacs.Columns[PCsAndMacsGridSortColumn], PCsAndMacsGridSortOrder);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully populate PC/MAC grid", ex);
            }
        }

        private void ClearPCsAndMACsEditPanel()
        {
            // Remove all selected information from each of the text boxes in PC / MAC edit panel.
            //   (Start with a clean slate)

            ClearPanelTextBoxes(PnlPCUser);
            ClearPanelTextBoxes(PnlPCHardware);

            ClearCombo(CmbPCType);
            ClearCombo(CmbPCDepartment);
            ClearCombo(CmbPCTemplate);

            RadPCActive.Checked = true;
            SetButtonState(CmdSavePC, false, Color.Black, Color.White);
        }

#endregion

#region ----PCs and MACs Helper Functions

        private void SavePCsAndMacsRecord(PCsAndMACsEditClass pCsMacs)
        {
            try
            {
                // Qualification tests:  
                //   SystemID and PC Type may not be null.
                if (QualifyPCsAndMacsValues(pCsMacs))
                {
                    // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.

                    if (pCsMacs.Status.IsNew)
                    {
                        if (!IsDuplicateSystemID(TxtPCSystemID.Text))
                        {
                            try
                            {
                                SqlParameter[] InsertParams = new SqlParameter[2];
                                InsertParams[0] = new SqlParameter("@pvchrpc_type", CmbPCType.Text);
                                InsertParams[1] = new SqlParameter("@pvchrsystem_id", TxtPCSystemID.Text);
                                SqlDataReader rdr = SQLQuery("Proc_Insert_PCsAndMACs", InsertParams);
                                using (rdr)
                                {
                                    rdr.Read();
                                    pCsMacs.Status.FieldList[0]["pcmac_id"] = SQLGetInt(rdr, "pcmac_id");
                                }

                            }
                            catch (Exception ex)
                            {
                                BroadcastWarning("ERROR trying to create new PC / MAC record", ex);
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("System ID " + TxtPCSystemID.Text + " is already in the database.  You cannot save a duplicate system ID.", "DUPLICATE SYSTEM ID", MessageBoxButtons.OK);
                            return;
                        }
                    }

                    // Update the selected PC/MAC record
                    SqlParameter[] PCParams = new SqlParameter[36];
                    PCParams[0] = new SqlParameter("@pvintpcmac_id", pCsMacs.Status.FieldList[0]["pcmac_id"]);
                    PCParams[1] = new SqlParameter("@pvchrpc_type", CmbPCType.Text);
                    PCParams[2] = new SqlParameter("@pvchrsystem_id", TxtPCSystemID.Text);
                    PCParams[3] = new SqlParameter("@pvchrusername", TxtPCUsername.Text);
                    PCParams[4] = new SqlParameter("@pvchruser_department", CmbPCDepartment.Text);
                    PCParams[5] = new SqlParameter("@pvintpc_template_id", PCsAndMacsRecord.Status.FieldList[0]["pc_template_id"]);
                    PCParams[6] = new SqlParameter("@pvchrsystem_tag", TxtPCSystemTag.Text);
                    PCParams[7] = new SqlParameter("@pvchrsystem_serial_number", TxtPCSystemSN.Text);
                    PCParams[8] = new SqlParameter("@pvchrmonitor_tag", TxtPCMonitorTag.Text);
                    PCParams[9] = new SqlParameter("@pvchrmonitor_serial_number", TxtPCMonitorSN.Text);
                    PCParams[10] = new SqlParameter("@pvchrkeyboard_tag", TxtPCKeyboardTag.Text);
                    PCParams[11] = new SqlParameter("@pvchrkeyboard_serial_number", TxtPCKeyboardSN.Text);
                    PCParams[12] = new SqlParameter("@pvchrmacaddress", TxtPCMACAddress.Text);
                    PCParams[13] = new SqlParameter("@pvchrmanufacturer", TxtPCManufacturer.Text);
                    PCParams[14] = new SqlParameter("@pvchrcabinet", TxtPCCabinet.Text);
                    PCParams[15] = new SqlParameter("@pvchrmodel", TxtPCModel.Text);
                    PCParams[16] = new SqlParameter("@pvchrprocessor", TxtPCProcessor.Text);
                    PCParams[17] = new SqlParameter("@pvchrram", TxtPCRAM.Text);
                    PCParams[18] = new SqlParameter("@pvchrmotherboards", TxtPCMotherboard.Text);
                    PCParams[19] = new SqlParameter("@pvchrmonitor1", TxtPCMonitor1.Text);
                    PCParams[20] = new SqlParameter("@pvchrmonitor2", TxtPCMonitor2.Text);
                    PCParams[21] = new SqlParameter("@pvchrhard_drive1", TxtPCHardDrive1.Text);
                    PCParams[22] = new SqlParameter("@pvchrhard_drive2", TxtPCHardDrive2.Text);
                    PCParams[23] = new SqlParameter("@pvchrmiscellaneous_drive", TxtPCMiscDrive.Text);
                    PCParams[24] = new SqlParameter("@pvchrnic", TxtPCNIC.Text);
                    PCParams[25] = new SqlParameter("@pvchrcd_device", TxtPCCDDevice.Text);
                    PCParams[26] = new SqlParameter("@pvchrvideo_card", TxtPCVideoCard.Text);
                    PCParams[27] = new SqlParameter("@pvchrsound_card", TxtPCSoundCard.Text);
                    PCParams[28] = new SqlParameter("@pvchrmiscellaneous_card", TxtPCMiscCard.Text);
                    PCParams[29] = new SqlParameter("@pvchrspeakers", TxtPCSpeakers.Text);
                    PCParams[30] = new SqlParameter("@pvchrkeyboard", TxtPCKeyboard.Text);
                    PCParams[31] = new SqlParameter("@pvchrmouse", TxtPCMouse.Text);
                    PCParams[32] = new SqlParameter("@pvchrcomments", TxtPCcomments.Text);
                    PCParams[33] = new SqlParameter("@pvbitactive_flag", RadPCActive.Checked ? 1 : 0);
                    PCParams[34] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                    PCParams[35] = new SqlParameter("@pvchrmodifiedby", UserInfo.Username);

                    SQLQuery("Proc_Update_PCsAndMACs", PCParams);

                    // Refresh the grid
                    ClearPCsAndMACsEditPanel();
                    PopulatePCsAndMacsGrid();

                    SetButtonState(CmdSavePC, false, Color.Black, Color.White);
                    CmbPCTemplate.ForeColor = Color.Black;
                    CmbPCType.ForeColor = Color.Black;
                    PCsAndMacsRecord.Status.IsSaved = true;
                    PCsAndMacsRecord.Status.IsNew = false;

                    // De-render the edit panels
                    PnlPCHardware.Visible = false;
                    PnlPCUser.Visible = false;

                }
                else
                {
                    MessageBox.Show("System ID and the PC Type MUST be specified prior to saving this record.  System ID must also be unique.", "MISSING SYSTEM ID AND PC TYPE", MessageBoxButtons.OK);
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to save new PC / MAC record", ex);
            }
        }

        private bool IsDuplicateSystemID(string systemid)
        {
            // Check if the passed-in system ID is already in the PCsAndMacs table.  Return true if it is.
            string SelectSTR = "SELECT * FROM PCsAndMACs WHERE system_id = '" + systemid + "'";
            using (SqlDataReader rdr = SQLQuery(SelectSTR, CommandType.Text))
            {
                return (rdr.HasRows);
            }
        }

        private bool QualifyPCsAndMacsValues(PCsAndMACsEditClass pCsAndMacsRecord)
        {
            // Qualification tests:  
            //   SystemID and PC Type may not be null.
            return ((TxtPCSystemID.MaxLength > 0) && (CmbPCType.Text.Length > 0));
        }
#endregion

#region ----PCs and MACs Events

        private void CmbPCTemplate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (IsInitializing) return;

            // Render the PC panel, and populate it with the defaults values from the selected template
            PnlPCHardware.Visible = true;

            // Get the record index and record corresponding to the selected template
            try
            {
                string SelectStr = "SELECT hardware_templates_id FROM TemplatesHardware WHERE template_name = '" + CmbPCTemplate.Text + "' ";
                using (SqlDataReader rdr = SQLQuery(SelectStr, CommandType.Text))
                {
                    if (rdr.HasRows)
                    {
                        rdr.Read();
                        int templateid = SQLGetInt(rdr, "hardware_templates_id");
                        PCsAndMacsRecord.Status.FieldList[0]["pc_template_id"] = templateid;

                        SqlParameter[] TemplateParams = new SqlParameter[1];
                        TemplateParams[0] = new SqlParameter("@pvintTemplateID", templateid);
                        List<Dictionary<string, object>> hwtemplate = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Hardware_Template", TemplateParams);
                        Dictionary<string, object> hwr = hwtemplate[0];

                        // Render the (new) Template values!
                        SafeText(TxtPCManufacturer, hwr, "Manufacturer");
                        SafeText(TxtPCCabinet, hwr, "Cabinet");
                        SafeText(TxtPCModel, hwr, "Model");
                        SafeText(TxtPCProcessor, hwr, "Processor");
                        SafeText(TxtPCRAM, hwr, "RAM");
                        SafeText(TxtPCMotherboard, hwr, "Motherboard");
                        SafeText(TxtPCMonitor1, hwr, "Monitor1");
                        SafeText(TxtPCMonitor2, hwr, "Monitor2");
                        SafeText(TxtPCHardDrive1, hwr, "Drive1");
                        SafeText(TxtPCHardDrive2, hwr, "Drive2");
                        SafeText(TxtPCMiscDrive, hwr, "MiscDrive");
                        SafeText(TxtPCNIC, hwr, "NIC");
                        SafeText(TxtPCCDDevice, hwr, "CD_Device");
                        SafeText(TxtPCVideoCard, hwr, "Video_Card");
                        SafeText(TxtPCSoundCard, hwr, "Sound_Card");
                        SafeText(TxtPCMiscCard, hwr, "Miscellaneous");
                        SafeText(TxtPCSpeakers, hwr, "Speakers");
                        SafeText(TxtPCKeyboard, hwr, "Keyboard");
                        SafeText(TxtPCMouse, hwr, "Mouse");
                        SafeText(TxtPCComments2, hwr, "comments");  // Note LOWERCASE!  Dictionaries are case-sensitive.
                    }
                }
                SetButtonState(CmdSavePC, true, Color.White, Color.Red);
                CmbPCTemplate.ForeColor = Color.Red;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to load / populate hardware template " + CmbPCTemplate.Text, ex);
            }
        }

        private void GrdPCsAndMacs_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction
                PCsAndMacsGridSortColumn = e.ColumnIndex;
                PCsAndMacsGridSortOrder = (GrdPCsAndMacs.Columns[PCsAndMacsGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            try
            {
                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (PCsAndMacsRecord != null)
                {
                    if (!PCsAndMacsRecord.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE PC / MAC RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SavePCsAndMacsRecord(PCsAndMacsRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the edit template. 
                PCsAndMacsRecord = new PCsAndMACsEditClass();
                PnlPCUser.Visible = true;
                PnlPCHardware.Visible = true;
                DisplayPCsAndMACsRecord(PCsAndMacsRecord, (int)GrdPCsAndMacs.Rows[row].Cells["pcmac_id"].Value);
                PCsAndMacsRecord.Status.IsNew = false;
                CmbPCTemplate.ForeColor = Color.Black;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully process PC/MAC grid entry", ex);
            }
        }

        private void PCsAndMacs_TextChanged(object sender, EventArgs e)
        {
            // This event is associated with ALL text and radio boxes in the PC / MAC tab's edit menu.
            //   See the PROPERTIES View for each of the test and radio boxes
            // On any change, enable the Save button
            if (!(PCsAndMacsRecord == null))
            {                
                //SetButtonState(CmdSavePC, false, Color.Black, Color.White);
                //PCsAndMacsRecord.Status.IsSaved = false;
                MarkAsDirty(sender, PCsAndMacsRecord.Status, CmdSavePC);
            }
        }

        private void CmdSavePC_Click(object sender, EventArgs e)
        {
            SavePCsAndMacsRecord(PCsAndMacsRecord);
        }

        private void CmdNewPC_Click(object sender, EventArgs e)
        {
            // If a new template entry has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            try
            {
                if (PCsAndMacsRecord != null)
                {
                    if (!PCsAndMacsRecord.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE PC / MAC RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SavePCsAndMacsRecord(PCsAndMacsRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Create a new record
                PCsAndMacsRecord = new PCsAndMACsEditClass();

                // enable user edit window for populating the new record
                ClearPCsAndMACsEditPanel();

                // On New, render the PC / MAC panels
                PnlPCUser.Visible = true;
                PnlPCHardware.Visible = false; // This gets enabled when user selects the appropriate hardware template
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully create new PC/MAC record", ex);
            }
        }

        private void RadPCActive_CheckedChanged(object sender, EventArgs e)
        {
            SetButtonState(CmdSavePC, true, Color.White, Color.Red);
        }

        private void RadPCInactive_CheckedChanged(object sender, EventArgs e)
        {
            SetButtonState(CmdSavePC, true, Color.White, Color.Red);
        }

        private void RadPcsAndMacsActiveStatus_Changed(object sender, EventArgs e)
        {
            PopulatePCsAndMacsGrid();
        }

#endregion

#endregion

#region Servers

#region ----Servers Declarations

        ServerEditClass ServerRecord;
        bool ServersInitialized = false;
        ListSortDirection ServersGridSortOrder = ListSortDirection.Ascending;
        ListSortDirection ServersContactGridSortOrder = ListSortDirection.Ascending;
        ListSortDirection ServersNetworkGridSortOrder = ListSortDirection.Ascending;
        ListSortDirection ServersOwnerGridSortOrder = ListSortDirection.Ascending;
        ListSortDirection ServersRAIDGridSortOrder = ListSortDirection.Ascending;

        int ServersGridSortColumn = 1;
        int ServersContactGridSortColumn = 2;
        int ServersNetworkGridSortColumn = 2;
        int ServersOwnerGridSortColumn = 3;
        int ServersRAIDGridSortColumn = 2;

#endregion

#region ----Servers Initialization
        private void PopulateServersGrid()
        {
            // Render all panels invisible

            PnlServer.Visible = false;
            PnlServerContact.Visible = false;
            PnlServerNetwork.Visible = false;
            PnlServerOwner.Visible = false;
            PnlServerRAID.Visible = false;

            // Populate all list boxes on the Servers tab

            PopulateGrid(GrdServers, "Proc_Select_Servers", CommandType.StoredProcedure, RadServerActive);
            PopulateComboBox(CmbServerManufacturer, "lstManufacturers", "manufacturer", "manufacturers_id");
            PopulateComboBox(CmbServerModel, "lstModels", "model", "models_id");
            PopulateComboBox(CmbServerProcessor, "lstProcessors", "processor", "processors_id");
            PopulateComboBox(CmbServerRAM, "lstRAM", "ram", "RAM_id");
            PopulateComboBox(CmbServerContainerType, "lstContainerTypes", "container_type", "container_types_id");
            PopulateComboBox(CmbServerOwner, "lstOwners", "owner_name", "owners_id");

            SetButtonState(CmdSaveServer, false, Color.Black, Color.White);
            // (For the related grids we need to specify some column attributes, but only AFTER the main grid's server selection is made 
            //   (or a NEW server is created)

            GrdServers.Columns[0].Visible = false;   // Hide the Server ID
            GrdServers.Sort(GrdServers.Columns[ServersGridSortColumn], ServersGridSortOrder);
        }

#endregion

#region ----Servers Classes

        public class ServerEditClass
        {
            // A simple class containing all fields in the Servers table
            public ServerContactsEditClass Contacts;
            public ServerNetworksEditClass Networks;
            public ServerOwnersEditClass Owners;
            public ServerRAIDEditClass RAID;
            public DictionaryStatusClass Status;

            public ServerEditClass()
            {
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["servers_id"] = 0,
                    ["server_name"] = "",
                    ["manufacturers_id"] = 0,
                    ["model_id"] = 0,
                    ["processor_id"] = 0,
                    ["ram_id"] = 0,
                    ["cpu_quantity"] = 0,
                    ["hard_drive_quantity"] = 0,
                    ["rack_number"] = 0,
                    ["active_flag"] = true
                };
                Status.FieldList.Add(fields);

                // Add contacts, networks, owners and RAID lists

                Contacts = new ServerContactsEditClass();
                Networks = new ServerNetworksEditClass();
                Owners = new ServerOwnersEditClass();
                RAID = new ServerRAIDEditClass();
            }
        }

        public class ServerContactsEditClass
        {
            // A simple class containing all fields in the Servers table
            public DictionaryStatusClass Status;

            public ServerContactsEditClass()
            {
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["servers_contacts_id"] = 0,
                    ["contact_type"] = "",
                    ["active_flag"] = true
                };
                Status.FieldList.Add(fields);
            }
        }

        public class ServerNetworksEditClass
        {
            // A simple class containing all fields in the Servers table
            public DictionaryStatusClass Status;

            public ServerNetworksEditClass()
            {
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["servers_networks_id"] = 0,
                    ["card_number"] = 1,
                    ["active_flag"] = true
                };
                Status.FieldList.Add(fields);
            }
        }

        public class ServerOwnersEditClass
        {
            // A simple class containing all fields in the Servers table
            public DictionaryStatusClass Status;

            public ServerOwnersEditClass()
            {
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["servers_owners_id"] = 0,
                    ["servers_id"] = 0,
                    ["owners_id"] = 0,
                    ["active_flag"] = true
                };
                Status.FieldList.Add(fields);
            }
        }

        public class ServerRAIDEditClass
        {
            // A simple class containing all fields in the Servers table
            public DictionaryStatusClass Status;

            public ServerRAIDEditClass()
            {
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["servers_raid_id"] = 0,
                    ["servers_id"] = 0,
                    ["container_number"] = 1,
                    ["active_flag"] = true
                };
                Status.FieldList.Add(fields);
            }
        }

#endregion

#region ----Servers Data Display Functions

        private void DisplayServerRecord(ServerEditClass serverRecord, int recordID)
        {
            try
            {
                ClearServerEditPanel();

                // We have 5 different records to get:
                //   Server
                //   Server Contacts
                //   Server Networks
                //   Server Owners
                //   Server RAID
                // We'll do this one at a time

                // Selected Server
                SqlParameter[] ServerParams = new SqlParameter[1];
                ServerParams[0] = new SqlParameter("@pvintServerID", recordID);
                serverRecord.Status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Servers", ServerParams);
                Dictionary<string, object> svr = serverRecord.Status.FieldList[0];

                SafeText(TxtServerName, svr, "server_name");
                SafeText(TxtServerPurpose, svr, "purpose");
                SafeText(TxtServerDetailedDescription, svr, "detailed_description");
                SafeComboBox(CmbServerManufacturer, SQLGetValueStringFromID("manufacturer", "lstManufacturers", "manufacturers_id", svr["manufacturers_id"].ToString()));
                SafeText(TxtServerSerialNumber, svr, "serial_number");
                SafeComboBox(CmbServerModel, SQLGetValueStringFromID("model", "lstModels", "models_id", svr["model_id"].ToString()));
                SafeText(TxtServerPrimaryIPAddr, svr, "primary_ip_address");
                SafeText(TxtServerOperatingSystem, svr, "operating_system");
                SafeDateBox(DTServerBuildDate, svr, "build_date");
                SafeText(TxtServerRackNumber, svr, "rack_number");
                SafeText(TxtServerSlotNumber, svr, "slot_number");
                SafeText(TxtServerSequence, svr, "sequence");
                SafeText(TxtServerLocationOther, svr, "physical_location_other");
                SafeRadioBox(RadServerActive, RadServerInactive, svr, "active_flag");
                SafeText(TxtServerCPUQty, svr, "cpu_quantity");
                SafeComboBox(CmbServerProcessor, SQLGetValueStringFromID("processor", "lstProcessors", "processors_id", svr["processor_id"].ToString()));
                SafeText(TxtServerProcessorDescription, svr, "cpu_description");
                SafeComboBox(CmbServerRAM, SQLGetValueStringFromID("ram", "lstRAM", "RAM_id", svr["ram_id"].ToString()));
                SafeText(TxtServerDriveQty, svr, "hard_drive_quantity");
                SafeText(TxtServerDriveCapacity, svr, "hard_drive_capacity");
                SafeText(TxtServerDrivePhysicalSize, svr, "hard_drive_physical_size");
                SafeText(TxtServerDriveNotes, svr, "hard_drive_notes");
                SafeText(TxtServerPrimaryApplications, svr, "primary_applications");
                SafeText(TxtServerAppInterfaces, svr, "application_interfaces");
                SafeText(TxtServerAppServices, svr, "application_services");
                SafeDateBox(DTServerPurchaseDate, svr, "purchase_date");
                SafeText(TxtServerPurchasePrice, svr, "purchase_price");
                SafeDateBox(DTServerWarrantyStart, svr, "warranty_start");
                SafeDateBox(DTServerWarrantyEnd, svr, "warranty_end");
                SafeText(TxtServerPrimaryUsers, svr, "primary_users");
                SafeText(TxtServerGeneralConfigFileLocations, svr, "configuration_file_locations");
                SafeText(TxtServerGeneralDirectoryLocations, svr, "directory_locations");
                SafeText(TxtServerGeneralRemarks, svr, "remarks");
                SafeText(TxtServerGeneralRestrictions, svr, "restrictions");

                // Populate Server Contacts Grid (using the server ID as the link) - This may be an empty set
                // Populate Server Networks Grid (using the server ID as the link) - This may be an empty set
                // Populate Server RAID Grid (using the server ID as the link) - This may be an empty set

                PopulateServerSubGrid(GrdServerContact, "SELECT * FROM Servers_Contacts WHERE servers_id = " + recordID.ToString(), CmdServerSaveContact, RadServersContactsFilterActive);
                PopulateServerSubGrid(GrdServerNetwork, "SELECT * FROM Servers_Networks WHERE servers_id = " + recordID.ToString(), CmdServerSaveNetwork, RadServersNetworkFilterActive);
                PopulateServerSubGrid(GrdServerRaid, "SELECT * FROM Servers_RAID WHERE servers_id = " + recordID.ToString(), CmdServerSaveRAID, RadServersRaidFilterActive);

                // Special case: Populate the Owners Grid.  This has a many-to-many relationship between Servers and Owners so it requires special processing
                //   TBD Might want to make this a stored procedure as it will likely be used elsewhere
                string SelectStr =
                    "SELECT O.owners_id, S.servers_id, S.servers_owners_id, O.owner_name, S.active_flag, S.last_modified_by, S.modifieddate " +
                    "FROM lstOwners O " +
                    "INNER JOIN Servers_Owners S ON O.owners_id = S.owners_id " +
                    "WHERE S.servers_id = " + recordID;
                if (RadServersOwnersFilterActive.Checked)
                {
                    SelectStr += " AND S.active_flag = 1";
                }
                if (RadServersOwnersFilterInactive.Checked)
                {
                    SelectStr += " AND S.active_flag = 0";
                }
                PopulateServerSubGrid(GrdServerOwner, SelectStr, CmdServerSaveOwner, null);

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdSaveServer, false, Color.Black, Color.White);
                serverRecord.Status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR displaying Server records", ex);
            }
        }

        private void DisplayServerOwnerRecord(DictionaryStatusClass status, int ownerRecordID)
        {
            try
            {
                // TBD MAKE THIS A SPROC - To be consistent, this sproc should be named "Proc_Select_Servers_Owners
                string OwnerStr =
                    "SELECT O.owners_id, S.servers_id, S.servers_owners_id, O.owner_name, S.active_flag, S.last_modified_by, S.modifieddate " +
                    "FROM lstOwners O " +
                    "INNER JOIN Servers_Owners S ON O.owners_id = S.owners_id " +
                    "WHERE S.servers_owners_id = " + ownerRecordID;
                status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.Text, OwnerStr);
                Dictionary<string, object> sow = status.FieldList[0];

                SafeText(CmbServerOwner, sow, "owner_name");
                SafeRadioBox(RadServerOwnerActive, RadServerOwnerInactive, sow, "active_flag");

                // TBD MAKE THIS A SPROC - THIS IS IDENTICAL TO Proc_Select_Servers_Owners, which should be renamed Proc_Select_Servers_Owners_By_Server_ID
                string GridStr =
                    "SELECT O.owners_id, S.servers_id, S.servers_owners_id, O.owner_name, S.active_flag, S.last_modified_by, S.modifieddate " +
                    "FROM lstOwners O " +
                    "INNER JOIN Servers_Owners S ON O.owners_id = S.owners_id " +
                    "WHERE S.servers_id = " + status.FieldList[0]["servers_id"].ToString();
                PopulateServerSubGrid(GrdServerOwner, GridStr, CmdServerSaveOwner, RadServersOwnersFilterActive);

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdServerSaveOwner, false, Color.Black, Color.White);
                status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)

                // Disable the Owner selection as well, since this can ONLY activate/deactivate the owner, not change the name
                CmbServerOwner.Enabled = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully display Server Owner record", ex);
            }
        }

        private void DisplayServerRaidRecord(DictionaryStatusClass status, int raidRecordID)
        {
            try
            {
                SqlParameter[] RaidParams = new SqlParameter[1];
                RaidParams[0] = new SqlParameter("@pvintRAIDID", raidRecordID);
                status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Servers_RAID", RaidParams);
                Dictionary<string, object> sra = status.FieldList[0];

                SafeText(CmbServerRaidType, sra, "raid_type");
                SafeText(CmbServerContactType, sra, "container_type");
                SafeText(TxtServerContainerNumber, sra, "container_number");
                SafeText(TxtServerContainerSize, sra, "container_size");
                SafeText(TxtServerRAIDNotes, sra, "raid_notes");
                SafeRadioBox(RadServerRAIDActive, RadServerRAIDInactive, sra, "active_flag");

                PopulateServerSubGrid(GrdServerRaid, "SELECT * FROM Servers_RAID WHERE servers_id = " + status.FieldList[0]["servers_id"].ToString(), CmdServerSaveRAID, RadServersRaidFilterActive);

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdServerSaveRAID, false, Color.Black, Color.White);
                status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully display Server RAID record", ex);
            }
        }

        private void DisplayServerNetworkRecord(DictionaryStatusClass status, int networkRecordID)
        {
            try
            {
                SqlParameter[] NetworkParams = new SqlParameter[1];
                NetworkParams[0] = new SqlParameter("@pvintNetworkID", networkRecordID);
                status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Servers_Networks", NetworkParams);
                Dictionary<string, object> sne = status.FieldList[0];

                SafeText(TxtServerNetworkCardNumber, sne, "card_number");
                SafeText(TxtServerNetworkIPAddress, sne, "ipaddress");
                SafeText(TxtServerNetworkMACAddress, sne, "macaddress");
                SafeText(TxtServerNetworkPatchLocation, sne, "patch_location");
                SafeText(TxtServerNetworkNIC, sne, "nic");
                SafeRadioBox(RadServerNetworkActive, RadServerNetworkInactive, sne, "active_flag");

                PopulateServerSubGrid(GrdServerNetwork, "SELECT * FROM Servers_Networks WHERE servers_id = " + status.FieldList[0]["servers_id"].ToString(), CmdServerSaveNetwork, RadServersNetworkFilterActive);

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdServerSaveNetwork, false, Color.Black, Color.White);
                status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully display Server Network record", ex);
            }
        }

        private void DisplayServerContactRecord(DictionaryStatusClass status, int contactRecordID)
        {
            try
            {
                SqlParameter[] ContactParams = new SqlParameter[1];
                ContactParams[0] = new SqlParameter("@pvintContactsID", contactRecordID);
                status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Servers_Contacts", ContactParams);
                Dictionary<string, object> sco = status.FieldList[0];

                SafeText(CmbServerContactType, sco, "contact_type");
                SafeText(TxtServerContact, sco, "contact");
                SafeText(TxtServerContactContractNumber, sco, "contract_number");
                SafeText(TxtServerContactPINNumber, sco, "pin_number");
                SafeText(TxtServerContactLicenseKey, sco, "license_key");
                SafeText(TxtServerContactSiteCodeID, sco, "site_code_id");
                SafeRadioBox(RadServerContactActive, RadServerContactInactive, sco, "active_flag");

                PopulateServerSubGrid(GrdServerNetwork, "SELECT * FROM Servers_Contacts WHERE servers_id = " + status.FieldList[0]["servers_id"].ToString(), CmdServerSaveContact, RadServersContactsFilterActive);

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdServerSaveContact, false, Color.Black, Color.White);
                status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully display Server Contact record", ex);
            }
        }

        private void PopulateServerSubGrid(DataGridView subGrid, string SelectStr, Button saveButton, RadioButton activeRadioButton)
        {
            // Generic routine to populate all server sub-grids.  These grids are constructed so that
            //   the first two fields are always the target table's record ID and servers_id, respectively.
            //   Both of these fields will be hidden from view.

            try
            {
                PopulateGrid(subGrid, SelectStr, CommandType.Text, activeRadioButton);
                SetButtonState(saveButton, false, Color.Black, Color.White);
                // For this grid we need to specify column attributes
                subGrid.Columns[0].Visible = false;  // Hide table record id
                subGrid.Columns[1].Visible = false;  // Hide servers_id
                subGrid.Sort(subGrid.Columns[2], ListSortDirection.Ascending); // default sort to the 3rd field, ascending
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR: Unable to successfully populate Server subgrid " + subGrid.Name, ex);
            }
        }

        private void ClearServerEditPanel()
        {
            // Remove all selected information from each of the text boxes in Server edit panel.
            //   (Start with a clean slate)

            ClearPanelTextBoxes(PnlServer);
            ClearCombo(CmbServerManufacturer);
            ClearCombo(CmbServerModel);

            DTServerBuildDate.Value = DateTime.Today;  // TBD Change forecolor to black?????
            RadServerActive.Checked = true;

            ClearPanelTextBoxes(PnlServerRAID);

            ClearCombo(CmbServerProcessor);
            ClearCombo(CmbServerRAM);
            ClearServerRaidEditPanel();
            ClearServerNetworkEditPanel();
            ClearServerContactsEditPanel();
            ClearServerOwnersEditPanel();

            DTServerPurchaseDate.Value = DateTime.Today;
            DTServerWarrantyStart.Value = DateTime.Today;
            DTServerWarrantyEnd.Value = DateTime.Today;

            LblBuildDate.ForeColor = Color.Black; // Sustitutes for the fact that the forecolor on the DateTimePicker control can't be changed from basic black.

            // Clear everthing on the subtabs

            TxtServerCPUQty.Text = "";
            TxtServerProcessorDescription.Text = "";
            TxtServerDriveQty.Text = "";
            TxtServerDriveCapacity.Text = "";
            TxtServerDrivePhysicalSize.Text = "";
            TxtServerDriveNotes.Text = "";

            TxtServerPrimaryApplications.Text = "";
            TxtServerAppInterfaces.Text = "";
            TxtServerAppServices.Text = "";

            TxtServerPurchasePrice.Text = "";
            TxtServerPrimaryUsers.Text = "";
            TxtServerGeneralConfigFileLocations.Text = "";
            TxtServerGeneralDirectoryLocations.Text = "";
            TxtServerGeneralRemarks.Text = "";
            TxtServerGeneralRestrictions.Text = "";
        }

        private void ClearServerOwnersEditPanel()
        {
            ClearCombo(CmbServerOwner);
            RadServerOwnerActive.Checked = true;
        }

        private void ClearServerContactsEditPanel()
        {
            ClearPanelTextBoxes(PnlServerContact);
            ClearCombo(CmbServerContactType);
            RadServerContactActive.Checked = true;
        }

        private void ClearServerNetworkEditPanel()
        {
            ClearPanelTextBoxes(PnlServerNetwork);
            RadServerNetworkActive.Checked = true;
        }

        private void ClearServerRaidEditPanel()
        {
            ClearPanelTextBoxes(PnlServerRAID);
            ClearCombo(CmbServerRaidType);
            ClearCombo(CmbServerContactType);
            RadServerRAIDActive.Checked = true;
        }
#endregion

#region ----Servers Helper Functions

        private void SaveServerRecord(ServerEditClass serverRecord)
        {

            try
            {
                // Qualification tests:  
                //   ServerName may not be null and must be unique
                //   Manufacturers ID, Model ID, Processor ID and RAM ID must all be nonzero

                string strMissing = "";
                strMissing += (TxtServerName.Text.Length <= 0 ? "SERVER NAME   " : "");
                strMissing += ((int)serverRecord.Status.FieldList[0]["manufacturers_id"] == 0 ? "MANUFACTURER   " : "");
                strMissing += ((int)serverRecord.Status.FieldList[0]["model_id"] == 0 ? "MODEL   " : "");
                strMissing += ((int)serverRecord.Status.FieldList[0]["processor_id"] == 0 ? "PROCESSOR   " : "");
                strMissing += ((int)serverRecord.Status.FieldList[0]["ram_id"] == 0 ? "RAM   " : "");

                if (strMissing.Length > 0)
                {
                    MessageBox.Show("The following items must ALL be specified prior to saving this record (Server Name, Manufacturer, Model, Processor and RAM).", "MISSING " + strMissing, MessageBoxButtons.OK);
                    return;  // One or more items is missing.
                }

                // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.
                if (serverRecord.Status.IsNew)
                {
                    try
                    {
                        // As a new record, servername better not already be in the table!
                        string SqlString = "SELECT server_name FROM Servers WHERE server_name = '" + TxtServerName.Text + "'";
                        using (SqlDataReader rdr = SQLQuery(SqlString, CommandType.Text))
                        {
                            if (rdr.HasRows)
                            {
                                MessageBox.Show("There is already a server with this name. Server name MUST be unique prior to saving this record.", "MISSING SERVER NAME", MessageBoxButtons.OK);
                                return;
                            }
                        }

                        // Create a new record
                        SqlParameter[] InsertParams = new SqlParameter[1];
                        InsertParams[0] = new SqlParameter("@pvchrServerName", TxtServerName.Text);
                        using (SqlDataReader rdr = SQLQuery("Proc_Insert_Servers", InsertParams))
                        {
                            rdr.Read();
                            serverRecord.Status.FieldList[0]["servers_id"] = SQLGetInt(rdr, "servers_id");
                        }
                    }
                    catch (Exception ex)
                    {
                        BroadcastWarning("ERROR trying to create new Server record", ex);
                        return;
                    }
                }

                // Now update the record with everything from the Server editor

                SqlParameter[] ServerParams = new SqlParameter[37];
                ServerParams[0] = new SqlParameter("@pvintserverid", serverRecord.Status.FieldList[0]["servers_id"]);
                ServerParams[1] = new SqlParameter("@pvchrserver_name", TxtServerName.Text);
                ServerParams[2] = new SqlParameter("@pvchrserial_number", TxtServerSerialNumber.Text);
                ServerParams[3] = new SqlParameter("@pvchrpurpose", TxtServerPurpose.Text);
                ServerParams[4] = new SqlParameter("@pvdatpurchase_date", DTServerPurchaseDate.Value.ToString());
                ServerParams[5] = new SqlParameter("@pvfltpurchase_price", TxtServerPurchasePrice.Text);
                ServerParams[6] = new SqlParameter("@pvdatbuild_date", DTServerBuildDate.Value.ToString());
                ServerParams[7] = new SqlParameter("@pvdatwarranty_start", DTServerWarrantyStart.Value.ToString());
                ServerParams[8] = new SqlParameter("@pvdatwarranty_end", DTServerWarrantyEnd.Value.ToString());
                ServerParams[9] = new SqlParameter("@pvchrprimary_users", TxtServerPrimaryUsers.Text);
                ServerParams[10] = new SqlParameter("@pvchrapplication_interfaces", TxtServerAppInterfaces.Text);
                ServerParams[11] = new SqlParameter("@pvchrapplication_services", TxtServerAppServices.Text);
                ServerParams[12] = new SqlParameter("@pvchrprimary_applications", TxtServerPrimaryApplications.Text);
                ServerParams[13] = new SqlParameter("@pvchrconfiguration_file_locations", TxtServerGeneralConfigFileLocations.Text);
                ServerParams[14] = new SqlParameter("@pvchrcpu_description", TxtServerProcessorDescription.Text);
                ServerParams[15] = new SqlParameter("@pvintcpu_quantity", TxtServerCPUQty.Text);
                ServerParams[16] = new SqlParameter("@pvchrdetailed_description", TxtServerDetailedDescription.Text);
                ServerParams[17] = new SqlParameter("@pvchrdirectory_locations", TxtServerGeneralDirectoryLocations.Text);
                ServerParams[18] = new SqlParameter("@pvinthard_drive_quantity", TxtServerDriveQty.Text);
                ServerParams[19] = new SqlParameter("@pvchrhard_drive_capacity", TxtServerDriveCapacity.Text);
                ServerParams[20] = new SqlParameter("@pvchrhard_drive_notes", TxtServerDriveNotes.Text);
                ServerParams[21] = new SqlParameter("@pvchrhard_drive_physical_size", TxtServerDrivePhysicalSize.Text);
                ServerParams[22] = new SqlParameter("@pvintmanufacturers_id", serverRecord.Status.FieldList[0]["manufacturers_id"]); // TBD Confirm that the IDs get loaded into the server record AFTER subgrid update!!!!!!
                ServerParams[23] = new SqlParameter("@pvintmodel_id", serverRecord.Status.FieldList[0]["model_id"]);
                ServerParams[24] = new SqlParameter("@pvintprocessor_id", serverRecord.Status.FieldList[0]["processor_id"]);
                ServerParams[25] = new SqlParameter("@pvintram_id", serverRecord.Status.FieldList[0]["ram_id"]);
                ServerParams[26] = new SqlParameter("@pvchrprimary_ip_address", TxtServerPrimaryIPAddr.Text);
                ServerParams[27] = new SqlParameter("@pvchroperating_system", TxtServerOperatingSystem.Text);
                ServerParams[28] = new SqlParameter("@pvchrphysical_location_other", TxtServerLocationOther.Text);
                ServerParams[29] = new SqlParameter("@pvintrack_number", TxtServerRackNumber.Text);
                ServerParams[30] = new SqlParameter("@pvintslot_number", TxtServerSlotNumber.Text);
                ServerParams[31] = new SqlParameter("@pvchrsequence", TxtServerSequence.Text);
                ServerParams[32] = new SqlParameter("@pvchrremarks", TxtServerGeneralRemarks.Text);
                ServerParams[33] = new SqlParameter("@pvchrrestrictions", TxtServerGeneralRestrictions.Text);
                ServerParams[34] = new SqlParameter("@pvbitactive_flag", RadServerActive.Checked ? 1 : 0);
                ServerParams[35] = new SqlParameter("@pvchrlast_modified_by", UserInfo.Username);
                ServerParams[36] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                SQLQuery("Proc_Update_Servers", ServerParams);

                // Refresh the grid
                ClearServerEditPanel();
                PopulateServersGrid();

                SetButtonState(CmdSaveServer, false, Color.Black, Color.White);
                serverRecord.Status.IsSaved = true;
                serverRecord.Status.IsNew = false;

                // De-render the edit panel
                PnlServer.Visible = false;

            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to update Server record", ex);
            }
        }

        private void SaveServerRaidRecord(ServerEditClass serverRecord)
        {
            try
            {
                // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.
                DictionaryStatusClass StatusRecord = serverRecord.RAID.Status;
                if (StatusRecord.IsNew)
                {
                    try
                    {
                        // Qualifier:  ContainerNumber must be an integer >0 =
                        bool isokay = int.TryParse(TxtServerContainerNumber.Text, out int containernumber);
                        if ((!isokay) || (containernumber < 0))
                        {
                            MessageBox.Show("ERROR - Container Number must be set to an integer value > 0", "INVALID CONTAINER NUMBER", MessageBoxButtons.OK);
                            return;
                        }
                        // Create a new record
                        SqlParameter[] InsertParams = new SqlParameter[2];
                        InsertParams[0] = new SqlParameter("@pvintServerID", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                        InsertParams[1] = new SqlParameter("@pvintContainerNumber", TxtServerContainerNumber.Text);
                        using (SqlDataReader rdr = SQLQuery("Proc_Insert_Servers_RAID", InsertParams))
                        {
                            rdr.Read();
                            StatusRecord.FieldList[0]["servers_raid_id"] = SQLGetInt(rdr, "servers_raid_id");
                        }
                    }
                    catch (Exception ex)
                    {
                        BroadcastWarning("ERROR trying to create new RAID record", ex);
                        return;
                    }
                }

                // Now update the record with everything from the Server's RAID editor

                SqlParameter[] RaidParams = new SqlParameter[10];
                RaidParams[0] = new SqlParameter("@pvintserversraidid", StatusRecord.FieldList[0]["servers_raid_id"].ToString());
                RaidParams[1] = new SqlParameter("@pvintserversid", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                RaidParams[2] = new SqlParameter("@pvchrraidtype", CmbServerRaidType.Text);
                RaidParams[3] = new SqlParameter("@pvchrcontainertype", CmbServerContainerType.Text);
                RaidParams[4] = new SqlParameter("@pvintcontainernumber", TxtServerContainerNumber.Text);
                RaidParams[5] = new SqlParameter("@pvchrcontainersize", TxtServerContainerSize.Text);
                RaidParams[6] = new SqlParameter("@pvchrraidnotes", TxtServerRAIDNotes.Text);
                RaidParams[7] = new SqlParameter("@pvbitactiveflag", RadServerRAIDActive.Checked ? 1 : 0);
                RaidParams[8] = new SqlParameter("@pvchrlastmodifiedby", UserInfo.Username);
                RaidParams[9] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                SQLQuery("Proc_Update_Servers_RAID", RaidParams);

                // Refresh the grid
                ClearServerRaidEditPanel();
                PopulateServerSubGrid(GrdServerRaid, "SELECT * FROM Servers_RAID WHERE servers_id = " + serverRecord.Status.FieldList[0]["servers_id"].ToString(), CmdServerSaveRAID, RadServersRaidFilterActive);
                GrdServerRaid.Sort(GrdServerRaid.Columns[ServersRAIDGridSortColumn], ServersRAIDGridSortOrder);

                SetButtonState(CmdServerSaveRAID, false, Color.Black, Color.White);
                StatusRecord.IsSaved = true;
                StatusRecord.IsNew = false;

                // De-render the edit panel
                PnlServerRAID.Visible = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to update RAID record", ex);
            }
        }

        private void SaveServerNetworkRecord(ServerEditClass serverRecord)
        {
            try
            {
                // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.
                DictionaryStatusClass StatusRecord = serverRecord.Networks.Status;
                if (StatusRecord.IsNew)
                {
                    try
                    {
                        // Qualifier:  Card Number must be an integer > 0
                        bool isokay = int.TryParse(TxtServerNetworkCardNumber.Text, out int cardnumber);
                        if ((!isokay) || (cardnumber < 0))
                        {
                            MessageBox.Show("ERROR - Card Number must be set to an integer value > 0", "INVALID CARD NUMBER", MessageBoxButtons.OK);
                            return;
                        }

                        // Create a new record
                        SqlParameter[] InsertParams = new SqlParameter[2];
                        InsertParams[0] = new SqlParameter("@pvintServerID", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                        InsertParams[1] = new SqlParameter("@pvintCardNumber", TxtServerNetworkCardNumber.Text);
                        using (SqlDataReader rdr = SQLQuery("Proc_Insert_Servers_Networks", InsertParams))
                        {
                            rdr.Read();
                            StatusRecord.FieldList[0]["servers_networks_id"] = SQLGetInt(rdr, "servers_networks_id");
                        }
                    }
                    catch (Exception ex)
                    {
                        BroadcastWarning("ERROR trying to create new Networks record", ex);
                        return;
                    }
                }

                // Now update the record with everything from the Server's Network editor

                SqlParameter[] NetworkParams = new SqlParameter[10];
                NetworkParams[0] = new SqlParameter("@pvintserversnetworksid", StatusRecord.FieldList[0]["servers_networks_id"].ToString());
                NetworkParams[1] = new SqlParameter("@pvintserversid", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                NetworkParams[2] = new SqlParameter("@pvintcardnumber", TxtServerNetworkCardNumber.Text);
                NetworkParams[3] = new SqlParameter("@pvchrnic", TxtServerNetworkNIC.Text);
                NetworkParams[4] = new SqlParameter("@pvchrmacaddress", TxtServerNetworkMACAddress.Text);
                NetworkParams[5] = new SqlParameter("@pvchripaddress", TxtServerNetworkIPAddress.Text);
                NetworkParams[6] = new SqlParameter("@pvchrpatchlocation", TxtServerNetworkPatchLocation.Text);
                NetworkParams[7] = new SqlParameter("@pvbitactiveflag", RadServerNetworkActive.Checked ? 1 : 0);
                NetworkParams[8] = new SqlParameter("@pvchrlastmodifiedby", UserInfo.Username);
                NetworkParams[9] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                SQLQuery("Proc_Update_Servers_Networks", NetworkParams);

                // Refresh the grid
                ClearServerNetworkEditPanel();
                PopulateServerSubGrid(GrdServerNetwork, "SELECT * FROM Servers_Networks WHERE servers_id = " + serverRecord.Status.FieldList[0]["servers_id"].ToString(), CmdServerSaveNetwork, RadServersNetworkFilterActive);
                GrdServerNetwork.Sort(GrdServerNetwork.Columns[ServersNetworkGridSortColumn], ServersNetworkGridSortOrder);

                SetButtonState(CmdServerSaveNetwork, false, Color.Black, Color.White);
                StatusRecord.IsSaved = true;
                StatusRecord.IsNew = false;

                // De-render the edit panel
                PnlServerNetwork.Visible = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to update Network record", ex);
            }
        }

        private void SaveServerContactRecord(ServerEditClass serverRecord)
        {
            try
            {
                // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.
                DictionaryStatusClass StatusRecord = serverRecord.Contacts.Status;
                if (StatusRecord.IsNew)
                {
                    try
                    {
                        // Qualifier:  Contact Type may not be null
                        if (CmbServerContactType.Text.Length <= 0)
                        {
                            MessageBox.Show("ERROR - Contact Type must be entered before saving", "INVALID CONTACT TYPE", MessageBoxButtons.OK);
                            return;
                        }

                        // Create a new record
                        SqlParameter[] InsertParams = new SqlParameter[2];
                        InsertParams[0] = new SqlParameter("@pvintServerID", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                        InsertParams[1] = new SqlParameter("@pvchrContactType", CmbServerContactType.Text);
                        using (SqlDataReader rdr = SQLQuery("Proc_Insert_Servers_Contacts", InsertParams))
                        {
                            rdr.Read();
                            StatusRecord.FieldList[0]["servers_contacts_id"] = SQLGetInt(rdr, "servers_contacts_id");
                        }
                    }
                    catch (Exception ex)
                    {
                        BroadcastWarning("ERROR trying to create new Contact record", ex);
                        return;
                    }
                }

                // Now update the record with everything from the Server's Network editor

                SqlParameter[] ContactParams = new SqlParameter[11];
                ContactParams[0] = new SqlParameter("@pvintserverscontactsid", StatusRecord.FieldList[0]["servers_contacts_id"].ToString());
                ContactParams[1] = new SqlParameter("@pvintserversid", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                ContactParams[2] = new SqlParameter("@pvchrcontact", TxtServerContact.Text);
                ContactParams[3] = new SqlParameter("@pvchrcontacttype", CmbServerContactType.Text);
                ContactParams[4] = new SqlParameter("@pvchrcontractnumber", TxtServerContactContractNumber.Text);
                ContactParams[5] = new SqlParameter("@pvchrlicensekey", TxtServerContactLicenseKey.Text);
                ContactParams[6] = new SqlParameter("@pvchrpinnumber", TxtServerContactPINNumber.Text);
                ContactParams[7] = new SqlParameter("@pvchrsitecodeid", TxtServerContactSiteCodeID.Text);
                ContactParams[8] = new SqlParameter("@pvbitactiveflag", RadServerContactActive.Checked ? 1 : 0);
                ContactParams[9] = new SqlParameter("@pvchrlastmodifiedby", UserInfo.Username);
                ContactParams[10] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                SQLQuery("Proc_Update_Servers_Contacts", ContactParams);

                // Refresh the grid
                ClearServerContactsEditPanel();
                PopulateServerSubGrid(GrdServerContact, "SELECT * FROM Servers_Contacts WHERE servers_id = " + serverRecord.Status.FieldList[0]["servers_id"].ToString(), CmdServerSaveContact, RadServersContactsFilterActive);
                GrdServerContact.Sort(GrdServerContact.Columns[ServersContactGridSortColumn], ServersContactGridSortOrder);

                SetButtonState(CmdServerSaveContact, false, Color.Black, Color.White);
                StatusRecord.IsSaved = true;
                StatusRecord.IsNew = false;

                // De-render the edit panel
                PnlServerContact.Visible = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to update Contact record", ex);
            }
        }

        private void SaveServerOwnerRecord(ServerEditClass serverRecord)
        {
            try
            {
                // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.
                DictionaryStatusClass StatusRecord = serverRecord.Owners.Status;
                if (StatusRecord.IsNew)
                {
                    try
                    {
                        // Create a new record
                        SqlParameter[] InsertParams = new SqlParameter[2];
                        InsertParams[0] = new SqlParameter("@pvintServerID", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                        InsertParams[1] = new SqlParameter("@pvintOwnersID", StatusRecord.FieldList[0]["owners_id"].ToString());
                        using (SqlDataReader rdr = SQLQuery("Proc_Insert_Servers_Owners", InsertParams))
                        {
                            rdr.Read();
                            StatusRecord.FieldList[0]["servers_owners_id"] = SQLGetInt(rdr, "servers_owners_id");
                        }
                    }
                    catch (Exception ex)
                    {
                        BroadcastWarning("ERROR trying to create new Owner record", ex);
                        return;
                    }
                }

                // Now update the (new) record with everything from the Server's Network editor

                SqlParameter[] OwnerParams = new SqlParameter[6];
                OwnerParams[0] = new SqlParameter("@pvintserversownersid", StatusRecord.FieldList[0]["servers_owners_id"].ToString());
                OwnerParams[1] = new SqlParameter("@pvintserversid", serverRecord.Status.FieldList[0]["servers_id"].ToString());
                OwnerParams[2] = new SqlParameter("@pvintownersid", CmbServerOwner.SelectedValue.ToString());
                OwnerParams[3] = new SqlParameter("@pvbitactiveflag", RadServerOwnerActive.Checked ? 1 : 0);
                OwnerParams[4] = new SqlParameter("@pvchrlastmodifiedby", UserInfo.Username);
                OwnerParams[5] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                SQLQuery("Proc_Update_Servers_Owners", OwnerParams);

                // Refresh the grid
                ClearServerOwnersEditPanel();
                string SelectStr =
                    "SELECT O.owners_id, S.servers_id, S.servers_owners_id, O.owner_name, O.active_flag, O.last_modified_by " +
                    "FROM lstOwners O " +
                    "INNER JOIN Servers_Owners S ON O.owners_id = S.owners_id " +
                    "WHERE S.servers_id = " + serverRecord.Status.FieldList[0]["servers_id"].ToString();  // TBD Make this a sproc
                PopulateServerSubGrid(GrdServerOwner, SelectStr, CmdServerSaveOwner, RadServersOwnersFilterActive);
                GrdServerOwner.Columns[2].Visible = false; // Hide the Owner ID
                GrdServerOwner.Sort(GrdServerOwner.Columns[ServersOwnerGridSortColumn], ServersOwnerGridSortOrder);

                SetButtonState(CmdServerSaveOwner, false, Color.Black, Color.White);
                StatusRecord.IsSaved = true;
                StatusRecord.IsNew = false;

                // De-render the edit panel
                PnlServerOwner.Visible = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to update Owner record", ex);
            }
        }

        private void ConfirmServerComboboxEntry(ComboBox cmb, Button button, List<Dictionary<string, object>> fieldList)
        {
            try
            {
                if (cmb.ValueMember.Length <= 0) return;  // no value to process - generally occurs during initialization.

                // If this entry isn't on the list then add it
                bool recordadded = AddServerComboboxEntry(cmb);

                // If the combobox entry was accepted (added to the table or already a member of the table) then update its dictionary entry in the server record
                if (recordadded)
                {
                    SetButtonState(button, true, Color.White, Color.Red);
                    string keyname = "";
                    switch (cmb.Name)
                    {
                        case "CmbServerManufacturer": keyname = "manufacturers_id"; break;
                        case "CmbServerModel": keyname = "model_id"; break;
                        case "CmbServerProcessor": keyname = "processor_id"; break;
                        case "CmbServerRAM": keyname = "ram_id"; break;
                        case "CmbServerRaidType": keyname = "raid_type"; break;
                        case "CmbServerContainerType": keyname = "container_type"; break;
                        case "CmbServerContactType": keyname = "contact_type"; break;
                        case "CmbServerOwner": keyname = "owners_id"; break;
                        default:
                            BroadcastWarning("ERROR - Missing Server Combobox entry " + cmb.Name + " in the Server's combobox EntryComplete event", null);
                            SetButtonState(button, false, Color.Black, Color.White);
                            break;
                    }
                    fieldList[0][keyname] = cmb.SelectedValue;
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR while trying to confirm a combobox entry", ex);
            }
        }

#endregion

#region ----Servers Events

#region --------Servers Events New Button
        private void CmdNewServer_Click(object sender, EventArgs e)
        {
            // If a new Server record has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            try
            {
                if (ServerRecord != null)
                {
                    if ((!ServerRecord.Status.IsSaved) && (CmdSaveServer.Enabled))
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE SERVER RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Create a new record
                ServerRecord = new ServerEditClass();

                // enable user edit window for populating the new record
                ClearServerEditPanel();

                // On New, render the Server panels
                PnlServer.Visible = true;
                SetButtonState(CmdSaveServer, false, Color.Black, Color.White);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully save Server Record", ex);
            }
        }

        private void CmdServerNewNetwork_Click(object sender, EventArgs e)
        {
            // If a new Server Network record has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            try
            {
                DictionaryStatusClass status = ServerRecord.Networks.Status;
                if (status != null)
                {
                    if ((status.IsNew) && (!status.IsSaved))
                    {
                        DialogResult dlgResult = MessageBox.Show("The current Network record in the edit window has NOT been saved; do you want to save it first?", "SAVE NETWORK RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerNetworkRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // enable user edit window for populating the new record
                ClearServerNetworkEditPanel();

                // Create a new record
                ServerRecord.Networks = new ServerNetworksEditClass();

                // On New, render the Server panels
                PnlServerNetwork.Visible = true;
                SetButtonState(CmdServerSaveNetwork, false, Color.Black, Color.White);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully save Network Record", ex);
            }
        }

        private void CmdServerNewContact_Click(object sender, EventArgs e)
        {
            // If a new Server Network record has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            try
            {
                DictionaryStatusClass status = ServerRecord.Contacts.Status;
                if (status != null)
                {
                    if ((status.IsNew) && (!status.IsSaved))
                    {
                        DialogResult dlgResult = MessageBox.Show("The current Contacts record in the edit window has NOT been saved; do you want to save it first?", "SAVE CONTACT RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerContactRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // enable user edit window for populating the new record
                ClearServerContactsEditPanel();

                // Create a new record
                ServerRecord.Contacts = new ServerContactsEditClass();

                // On New, render the Server panels
                PnlServerContact.Visible = true;
                SetButtonState(CmdServerSaveContact, false, Color.Black, Color.White);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully save Contact Record", ex);
            }
        }

        private void CmdServerNewOwner_Click(object sender, EventArgs e)
        {
            // Owner records CANNOT be updated (as they are used elsewhere) and can only be New'd.  
            // If a new Server Owner record has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            try
            {
                DictionaryStatusClass status = ServerRecord.Owners.Status;
                if (status != null)
                {
                    if ((status.IsNew) && (!status.IsSaved))
                    {
                        DialogResult dlgResult = MessageBox.Show("The current Owner record in the edit window has NOT been saved; do you want to save it first?", "SAVE OWNER RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerOwnerRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // enable user edit window for populating the new record
                ClearServerOwnersEditPanel();

                // Create a new record
                ServerRecord.Owners = new ServerOwnersEditClass();

                // On New, render the Server panels
                PnlServerOwner.Visible = true;
                SetButtonState(CmdServerSaveOwner, false, Color.Black, Color.White);
                CmbServerOwner.Enabled = true;
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully save Owner Record", ex);
            }
        }

        private void CmdServerNewRAID_Click(object sender, EventArgs e)
        {
            // If a new Server RAID record has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            try
            {
                DictionaryStatusClass status = ServerRecord.RAID.Status;
                if (status != null)
                {
                    if ((status.IsNew) && (!status.IsSaved))
                    {
                        DialogResult dlgResult = MessageBox.Show("The current RAID record in the edit window has NOT been saved; do you want to save it first?", "SAVE RAID RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerRaidRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // enable user edit window for populating the new record
                ClearServerRaidEditPanel();

                // Create a new record
                ServerRecord.RAID = new ServerRAIDEditClass();

                // On New, render the Server panels
                PnlServerRAID.Visible = true;
                SetButtonState(CmdServerSaveRAID, false, Color.Black, Color.White);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Unable to successfully save RAID Record", ex);
            }
        }
#endregion

#region --------Servers Events Save Button
        private void CmdSaveServer_Click(object sender, EventArgs e)
        {
            // Save the server data
            SaveServerRecord(ServerRecord);
        }

        private void CmdServerSaveRAID_Click(object sender, EventArgs e)
        {
            SaveServerRaidRecord(ServerRecord);
        }

        private void CmdServerSaveNetwork_Click(object sender, EventArgs e)
        {
            SaveServerNetworkRecord(ServerRecord);
        }

        private void CmdServerSaveOwner_Click(object sender, EventArgs e)
        {
            SaveServerOwnerRecord(ServerRecord);
        }

        private void CmdServerSaveContact_Click(object sender, EventArgs e)
        {
            SaveServerContactRecord(ServerRecord);
        }
#endregion

#region --------Servers Events Cell Click 
        private void GrdServers_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction
                ServersGridSortColumn = e.ColumnIndex;
                ServersGridSortOrder = (GrdServers.Columns[ServersGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            try
            {
                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (ServerRecord != null)
                {
                    if ((!ServerRecord.Status.IsSaved) && (CmdSaveServer.Enabled))
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE SERVER RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the edit template. 
                ServerRecord = new ServerEditClass();
                PnlServer.Visible = true;
                DisplayServerRecord(ServerRecord, (int)GrdServers.Rows[row].Cells["servers_id"].Value);
                ServerRecord.Status.IsNew = false;
                CmbServerManufacturer.ForeColor = Color.Black;
                CmbServerModel.ForeColor = Color.Black;
                CmbServerProcessor.ForeColor = Color.Black;
                CmbServerRAM.ForeColor = Color.Black;

                LblBuildDate.ForeColor = Color.Black;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully select Grid entry", ex);
            }
        }

        private void GrdServerRaid_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction                
                ServersRAIDGridSortColumn = e.ColumnIndex;
                ServersRAIDGridSortOrder = (GrdServerRaid.Columns[ServersRAIDGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            try
            {
                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (ServerRecord.RAID.Status != null)
                {
                    DictionaryStatusClass StatusRecord = ServerRecord.RAID.Status;
                    if (!StatusRecord.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current RAID record in the edit window has NOT been saved; do you want to save it first?", "SAVE RAID RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerRaidRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the edit template. 
                ServerRecord.RAID = new ServerRAIDEditClass();
                DictionaryStatusClass NewStatusRecord = ServerRecord.RAID.Status;
                DisplayServerRaidRecord(NewStatusRecord, (int)GrdServerRaid.Rows[row].Cells["servers_raid_id"].Value);
                NewStatusRecord.IsNew = false;

                PnlServerRAID.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully select Grid entry", ex);
            }
        }

        private void GrdServerNetwork_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction                
                ServersNetworkGridSortColumn = e.ColumnIndex;
                ServersNetworkGridSortOrder = (GrdServerNetwork.Columns[ServersNetworkGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            try
            {
                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (ServerRecord.Networks.Status != null)
                {
                    DictionaryStatusClass StatusRecord = ServerRecord.Networks.Status;
                    if (!StatusRecord.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current Network record in the edit window has NOT been saved; do you want to save it first?", "SAVE NETWORK RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerNetworkRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the edit template. 
                ServerRecord.Networks = new ServerNetworksEditClass();
                DictionaryStatusClass NewStatusRecord = ServerRecord.Networks.Status;
                DisplayServerNetworkRecord(NewStatusRecord, (int)GrdServerNetwork.Rows[row].Cells["servers_networks_id"].Value);
                NewStatusRecord.IsNew = false;

                PnlServerNetwork.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully select Grid entry", ex);
            }
        }

        private void GrdServerContact_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction                
                ServersContactGridSortColumn = e.ColumnIndex;
                ServersContactGridSortOrder = (GrdServerContact.Columns[ServersContactGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            try
            {
                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (ServerRecord.Contacts.Status != null)
                {
                    DictionaryStatusClass StatusRecord = ServerRecord.Contacts.Status;
                    if (!StatusRecord.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current Contact record in the edit window has NOT been saved; do you want to save it first?", "SAVE CONTACT RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerContactRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the edit template. 
                ServerRecord.Contacts = new ServerContactsEditClass();
                DictionaryStatusClass NewStatusRecord = ServerRecord.Contacts.Status;
                DisplayServerContactRecord(NewStatusRecord, (int)GrdServerContact.Rows[row].Cells["servers_contacts_id"].Value);
                NewStatusRecord.IsNew = false;

                PnlServerContact.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully select Grid entry", ex);
            }
        }

        private void GrdServerOwner_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Onwer records CANNOT be updated here (because they are used elsewhere.
            //   They can ONLY be New'd.
            // This event handler logic is here only to provide a reminder that if you tried to update
            //   an owner's name here, it gets updated wherever it's being used - so don't do it!

            // UPDATE:  Oh Wait:  We CAN set the owner inactive so we have to allow this.

            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction                
                ServersOwnerGridSortColumn = e.ColumnIndex;
                ServersOwnerGridSortOrder = (GrdServerOwner.Columns[ServersOwnerGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            try
            {
                // If a new template entry has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (ServerRecord.Owners.Status != null)
                {
                    DictionaryStatusClass StatusRecord = ServerRecord.Owners.Status;
                    if (!StatusRecord.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current Owner record in the edit window has NOT been saved; do you want to save it first?", "SAVE OWNER RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveServerOwnerRecord(ServerRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the edit template. 
                ServerRecord.Owners = new ServerOwnersEditClass();
                DictionaryStatusClass NewStatusRecord = ServerRecord.Owners.Status;
                DisplayServerOwnerRecord(NewStatusRecord, (int)GrdServerOwner.Rows[row].Cells["servers_owners_id"].Value);
                NewStatusRecord.IsNew = false;

                PnlServerOwner.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully select Grid entry", ex);
            }
        }
#endregion

#region --------Servers Events Data Entry
        /// <summary>
        /// An event unique to ComboBoxUnlocked.  It fires when the user depresses the ENTER key or when the combobox loses focus.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Event_ServerEntryComplete(object sender, EventArgs e)
        {
            // This is the generic _SelectedIndexChanged event for
            //    CmbServerManufacturerer
            //    CmbServerModel
            //    CmbServerProcessor
            //    CmbServerRAM

            ConfirmServerComboboxEntry((ComboBox)sender, CmdSaveServer, ServerRecord.Status.FieldList);
            MarkAsDirty(sender, ServerRecord.Status, CmdSaveServer);

            // Save the record ID of whatever combobox selection event just fired.
            ComboBox c = (ComboBox)sender;
            switch (c.Name.ToLower())
            {
                case "cmbservermanufacturer":
                    ServerRecord.Status.FieldList[0]["manufacturers_id"] = c.SelectedValue;
                    break;
                case "cmbservermodel":
                    ServerRecord.Status.FieldList[0]["model_id"] = c.SelectedValue;
                    break;
                case "cmbserverprocessor":
                    ServerRecord.Status.FieldList[0]["processor_id"] = c.SelectedValue;
                    break;
                case "cmbserverram":
                    ServerRecord.Status.FieldList[0]["ram_id"] = c.SelectedValue;
                    break;
                default:
                    MessageBox.Show("ERROR - Undefined CombBox '" + c.Name + " during server Save operation");
                    break;
            }
        }

        private void Event_ServerRaidEntryComplete(object sender, EventArgs e)
        {
            // Generic _SelectedIndexChanged event for
            //    CmbServerRaidType
            //    CmbServerContainerType

            ConfirmServerComboboxEntry((ComboBox)sender, CmdServerSaveRAID, ServerRecord.RAID.Status.FieldList);
            MarkAsDirty(sender, ServerRecord.RAID.Status, CmdServerSaveRAID);
        }

        private void Event_ServerContactEntryComplete(object sender, EventArgs e)
        {
            // Generic _SelectedIndexChanged event for
            //    CmbServerContactType

            ConfirmServerComboboxEntry((ComboBox)sender, CmdServerSaveContact, ServerRecord.Contacts.Status.FieldList);
            MarkAsDirty(sender, ServerRecord.Contacts.Status, CmdServerSaveContact);
        }

        private void Event_ServerOwnerEntryComplete(object sender, EventArgs e)
        {
            // Generic _SelectedIndexChanged event for
            //    CmbServerOwner

            ConfirmServerComboboxEntry((ComboBox)sender, CmdServerSaveOwner, ServerRecord.Owners.Status.FieldList);
            MarkAsDirty(sender, ServerRecord.Owners.Status, CmdServerSaveOwner);
        }

        //Server Data Changed Events
        private void Servers_TextChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Status, CmdSaveServer);
        }

        private void Servers_RadioButtonChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Status, CmdSaveServer);
        }

        private void Servers_DateValueChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Status, CmdSaveServer); // Does not work for datetimepicker controls (Windows limitation)
            LblBuildDate.ForeColor = Color.Red;
        }

        //RAID Data Changed Events
        private void Servers_RaidTextChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.RAID.Status, CmdServerSaveRAID);
        }

        private void Servers_RaidRadioButtonChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.RAID.Status, CmdServerSaveRAID);
        }

        // Network Data Changed Events
        private void Servers_NetworkTextChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Networks.Status, CmdServerSaveNetwork);
        }

        private void Servers_NetworksRadioButtonChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Networks.Status, CmdServerSaveNetwork);
        }

        // Contact Data Changed Events
        private void Servers_ContactTextChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Contacts.Status, CmdServerSaveContact);
        }

        private void Servers_ContactRadioButtonChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Contacts.Status, CmdServerSaveContact);
        }

        // Owner Data Changed Events
        private void Servers_OwnerTextChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Owners.Status, CmdServerSaveOwner);
        }

        private void Servers_OwnerRadioButtonChanged(object sender, EventArgs e)
        {
            MarkAsDirty(sender, ServerRecord.Owners.Status, CmdServerSaveOwner);
        }
#endregion

#region --------Servers Events Status Changes

        private void RadServersActiveStatus_Changed(object sender, EventArgs e)
        {
            PopulateGrid(GrdServers, "Proc_Select_Servers", CommandType.StoredProcedure, RadServerActive);
            //MarkAsDirty(sender, ServerRecord.Status, CmdSaveServer);
        }

        private void RadServersRaidActiveStatus_Changed(object sender, EventArgs e)
        {
            try
            {
                bool IDOkay = int.TryParse(ServerRecord.Status.FieldList[0]["servers_id"].ToString(), out int recordid);
                if (IDOkay && recordid > 0)
                {
                    PopulateServerSubGrid(GrdServerRaid, "SELECT * FROM Servers_RAID WHERE servers_id = " + recordid.ToString(), CmdServerSaveRAID, RadServersRaidFilterActive);
                }
                //MarkAsDirty(sender, ServerRecord.RAID.Status, CmdServerSaveRAID);
            }
            catch (Exception ex)
            {
                BroadcastWarning("WARNING:  RAID Active Status Change event failed", ex);
            }
        }

        private void RadServersNetworksActiveStatus_Changed(object sender, EventArgs e)
        {
            try
            {
                bool IDOkay = int.TryParse(ServerRecord.Status.FieldList[0]["servers_id"].ToString(), out int recordid);
                if (IDOkay && recordid > 0)
                {
                    PopulateServerSubGrid(GrdServerNetwork, "SELECT * FROM Servers_Networks WHERE servers_id = " + recordid, CmdServerSaveNetwork, RadServersNetworkFilterActive);
                }
                //MarkAsDirty(sender, ServerRecord.Networks.Status, CmdServerSaveNetwork);
            }
            catch (Exception ex)
            {
                BroadcastWarning("WARNING:  Networks Active Status Change event failed", ex);
            }
        }

        private void RadServersContactsActiveStatus_Changed(object sender, EventArgs e)
        {
            try
            {
                bool IDOkay = int.TryParse(ServerRecord.Status.FieldList[0]["servers_id"].ToString(), out int recordid);
                if (IDOkay && recordid > 0)
                {
                    PopulateServerSubGrid(GrdServerContact, "SELECT * FROM Servers_Contacts WHERE servers_id = " + recordid, CmdServerSaveContact, RadServersContactsFilterActive);
                }
                //MarkAsDirty(sender, ServerRecord.Contacts.Status, CmdServerSaveContact);
            }
            catch (Exception ex)
            {
                BroadcastWarning("WARNING:  Contacts Active Status Change event failed", ex);
            }
        }

        private void RadServersOwnersActiveStatus_Changed(object sender, EventArgs e)
        {
            try
            {
                bool IDOkay = int.TryParse(ServerRecord.Status.FieldList[0]["servers_id"].ToString(), out int recordid);
                if (IDOkay && recordid > 0)
                {
                    string SelectStr =
                    "SELECT O.owners_id, S.servers_id, S.servers_owners_id, O.owner_name, S.active_flag, S.last_modified_by, S.modifieddate " +
                    "FROM lstOwners O " +
                    "INNER JOIN Servers_Owners S ON O.owners_id = S.owners_id " +
                    "WHERE S.servers_id = " + recordid;
                    if (RadServersOwnersFilterActive.Checked)
                    {
                        SelectStr += " AND S.active_flag = 1";
                    }
                    if (RadServersOwnersFilterInactive.Checked)
                    {
                        SelectStr += " AND S.active_flag = 0";
                    }
                    PopulateServerSubGrid(GrdServerOwner, SelectStr, CmdServerSaveOwner, null);
                    //MarkAsDirty(sender, ServerRecord.Owners.Status, CmdServerSaveOwner);
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("WARNING:  Owners Active Status Change event failed", ex);
            }
        }

#endregion

#endregion

#endregion

#region Documentation

#region ----Documentation Declarations

        bool DocumentationInitialized = false;
        DocumentationClass DocumentationRecord;
        public int DocumentationGridSortColumn = 1;
        public int DocumentationVersionsGridSortColumn = 5;
        public ListSortDirection DocumentationGridSortOrder = ListSortDirection.Ascending;
        public ListSortDirection DocumentationVersionsGridSortOrder = ListSortDirection.Ascending;

#endregion

#region ----Documentation Initialization

        private void PopulateDocumentationGrid()
        {
            try
            {
                PopulateGrid(GrdDocumentation, "Proc_Select_Documentation", CommandType.StoredProcedure, RadDocumentationFilterActive);
                GrdDocumentation.Columns[0].Visible = false; // Hide the documentation_id field.
                GrdDocumentation.Sort(GrdDocumentation.Columns[DocumentationGridSortColumn], DocumentationGridSortOrder); // default sort to the 2nd field, ascending

                // Hide the documentation and version edit panels
                PnlDocumentation.Visible = false;
                PnlDocumentationVersion.Visible = false;
                GrdDocumentationVersion.Enabled = false;
                SetButtonState(CmdDocumentSave, false, Color.Black, Color.White);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying to populate the documentation grid", ex);
            }
        }

        private void PopulateDocumentationVersionsGrid(DocumentationClass documentRecord)
        {
            // ONLY Load the versions associated with the passed-in document record
            //   The embedded file field (a byte array) CANNOT be loaded onto the grid.

            try
            {
                // NOTE:  Because we want to select everything EXCEPT the embedded file field, filtered by documentation_id, 
                //   the PopulateGrid function is not designed to handle a stored procedure of this type so we use a SELECT query instead.
                string SelectStr = "SELECT [documentation_versions_id],[documentation_id],[file_extension],[description],[version_by],[version_date_time],[active_flag],[last_modified_by],[modifieddate] " +
                    "FROM Documentation_Versions WHERE documentation_id = " + documentRecord.Status.FieldList[0]["documentation_id"];
                PopulateGrid(GrdDocumentationVersion, SelectStr, CommandType.Text, RadDocumentationVersionFilterActive);
                GrdDocumentationVersion.Columns[0].Visible = false; // Hide the documentation_version_id field.
                GrdDocumentationVersion.Columns[1].Visible = false; // Hide the documentation_id field.
                GrdDocumentationVersion.Columns[2].Visible = false; // Hide the embedded file extension.

                GrdDocumentationVersion.Sort(GrdDocumentationVersion.Columns[DocumentationVersionsGridSortColumn], DocumentationVersionsGridSortOrder); // default sort to the version_id field, ascending

                // Hide the version edit panel
                PnlDocumentationVersion.Visible = false;
                SetButtonState(CmdDocumentVersionSave, false, Color.Black, Color.White);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying to populate the version grid", ex);
            }
        }

#endregion

#region ----Documentation Classes

        public class DocumentationClass
        {
            // A simple class containing all fields in the Documentation table
            public DocumentationVersionClass Versions;
            public DictionaryStatusClass Status;

            public DocumentationClass()
            {
                Versions = new DocumentationVersionClass();
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["documentation_id"] = 0,
                    ["subject"] = "",
                    ["active_flag"] = true
                };
                Status.FieldList.Add(fields);
            }
        }

        public class DocumentationVersionClass
        {
            // A simple class containing all fields in the DocumentationVersions table
            public DictionaryStatusClass Status;

            public DocumentationVersionClass()
            {
                Status = new DictionaryStatusClass();

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["documentation_versions_id"] = 0,
                    ["documentation_id"] = 0,
                    ["version_by"] = "",
                    ["version_date"] = "",
                    ["file_extension"] = "",
                    ["active_flag"] = true
                };
                Status.FieldList.Add(fields);
            }
        }

#endregion

#region ----Documentation Data Display Functions

        private void DisplayDocumentationVersionRecord(DocumentationClass documentationRecord, int recordID)
        {

            try
            {
                ClearDocumentVersionEditPanel();

                // Selected Document
                SqlParameter[] VersionParams = new SqlParameter[1];
                VersionParams[0] = new SqlParameter("@pvintDocumentationVersionsID", recordID);
                documentationRecord.Versions.Status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Documentation_Versions", VersionParams);
                Dictionary<string, object> ver = documentationRecord.Versions.Status.FieldList[0];

                SafeText(TxtDocumentationVersionDescription, ver, "description");
                SafeDateBox(DTDocumentationVersionDate, ver, "version_date_time");
                SafeRadioBox(RadDocumentationVersionActive, RadDocumentationVersionInactive, ver, "active_flag");

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdDocumentVersionSave, false, Color.Black, Color.White);
                documentationRecord.Versions.Status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR displaying Document Version records", ex);
            }
        }

        private void DisplayDocumentationRecord(DocumentationClass documentationRecord, int recordID)
        {
            try
            {
                ClearDocumentEditPanel();

                // We have 2 different records to get:
                //   Documentation
                //   Documentation Version
                // We'll do this one at a time

                // Selected Document
                SqlParameter[] DocumentParams = new SqlParameter[1];
                DocumentParams[0] = new SqlParameter("@pvintDocumentationID", recordID);
                documentationRecord.Status.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Documentation", DocumentParams);
                Dictionary<string, object> doc = documentationRecord.Status.FieldList[0];

                SafeText(TxTDocumentationSubject, doc, "subject");
                SafeText(TxtDocumentURL, doc, "url");
                SafeRadioBox(RadDocumentationActive, RadDocumentationInactive, doc, "active_flag");

                // Populate Documentatin Versions Grid (using the documentation ID as the link) - This may be an empty set

                PopulateDocumentationVersionsGrid(documentationRecord);

                // Disable the SAVE button until the user makes an edit.
                SetButtonState(CmdDocumentSave, false, Color.Black, Color.White);
                documentationRecord.Status.IsSaved = true; // Indicates that no edits to this record have occurred (yet)

                // If the URL is non-blank then enable the document view button
                CmdDocumentationViewDocument.Enabled = (TxtDocumentURL.Text.Trim().Length > 0);
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR displaying Document records", ex);
            }
        }

        private void ClearDocumentEditPanel()
        {
            // Remove all selected information from each of the text boxes in the Document edit panels.
            //   (Start with a clean slate)

            try
            {
                ClearPanelTextBoxes(PnlDocumentation);
                ClearPanelTextBoxes(PnlDocumentationVersion);

                DTDocumentationVersionDate.Value = DateTime.Today;
                RadDocumentationActive.Checked = true;
                RadDocumentationVersionActive.Checked = true;
                CmdDocumentationViewDocument.Enabled = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to clear the document edit panel", ex);
            }
        }

        private void ClearDocumentVersionEditPanel()
        {
            // Remove all selected information from each of the text boxes in the Document Version edit panels.
            //   (Start with a clean slate)

            try
            {
                ClearPanelTextBoxes(PnlDocumentationVersion);

                DTDocumentationVersionDate.Value = DateTime.Today;
                RadDocumentationVersionActive.Checked = true;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to clear the version edit panel", ex);
            }
        }

#endregion

#region ----Documentation Helper Functions

        private void SaveDocumentRecord(DocumentationClass documentationRecord)
        {
            try
            {
                // Qualification tests:
                //   Document URL must have been entered on the edit panel
                //   Subject may not be blank

                if (TxtDocumentURL.Text.Trim().Length <= 0)
                {
                    MessageBox.Show("Document name/URL MUST be specified prior to saving this record.", "MISSING DOCUMENT NAME", MessageBoxButtons.OK);
                    return;
                }

                if (TxTDocumentationSubject.Text.Trim().Length <= 0)
                {
                    MessageBox.Show("Document subject MUST be specified prior to saving this record.", "MISSING DOCUMENT SUBJECT", MessageBoxButtons.OK);
                    return;
                }

                // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.
                if (documentationRecord.Status.IsNew)
                {
                    try
                    {
                        // A new document will not have any associated document versions (it doesn't even yet have a record ID).  
                        // Create the new record now.
                        SqlParameter[] InsertParams = new SqlParameter[2];
                        InsertParams[0] = new SqlParameter("@pvchrURL", TxtDocumentURL.Text);
                        InsertParams[1] = new SqlParameter("@pvchrSubject", TxTDocumentationSubject.Text);
                        using (SqlDataReader rdr = SQLQuery("Proc_Insert_Documentation", InsertParams))
                        {
                            rdr.Read();
                            documentationRecord.Status.FieldList[0]["documentation_id"] = SQLGetInt(rdr, "documentation_id");
                        }

                        // We will also automatically save the version panel as well (in order to be sure that version 1 is created, even if it's totally empty).

                        SaveDocumentVersionRecord(documentationRecord, true);
                    }

                    catch (Exception ex)
                    {
                        BroadcastWarning("ERROR trying to create new Documentation record", ex);
                        return;
                    }
                }

                // Now update the record with everything from the Document editor
                SqlParameter[] DocumentParams = new SqlParameter[6];
                DocumentParams[0] = new SqlParameter("@pvintdocumentationid", documentationRecord.Status.FieldList[0]["documentation_id"]);
                DocumentParams[1] = new SqlParameter("@pvchrsubject", TxTDocumentationSubject.Text);
                DocumentParams[2] = new SqlParameter("@pvchrurl", TxtDocumentURL.Text);
                DocumentParams[3] = new SqlParameter("@pvbitactive_flag", (RadDocumentationActive.Checked ? 1 : 0));
                DocumentParams[4] = new SqlParameter("@pvchrlast_modified_by", UserInfo.Username);
                DocumentParams[5] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                SQLQuery("Proc_Update_Documentation", DocumentParams);

                // Refresh the grid
                ClearDocumentEditPanel();
                PopulateDocumentationGrid();

                SetButtonState(CmdDocumentSave, false, Color.Black, Color.White);
                documentationRecord.Status.IsSaved = true;
                documentationRecord.Status.IsNew = false;

                // De-render the edit panels
                PnlDocumentation.Visible = false;
                PnlDocumentationVersion.Visible = false;

            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to update Documentation record", ex);
            }
        }

        private void SaveDocumentVersionRecord(DocumentationClass documentationRecord, bool parentGenerated)
        {
            try
            {
                // Is this a new version record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.
                if (documentationRecord.Versions.Status.IsNew)
                {
                    try
                    {
                        SqlParameter[] InsertParams = new SqlParameter[1];
                        InsertParams[0] = new SqlParameter("@pvintDocumentationID", documentationRecord.Status.FieldList[0]["documentation_id"]);
                        using (SqlDataReader rdr = SQLQuery("Proc_Insert_Documentation_Versions", InsertParams))
                        {
                            rdr.Read();
                            documentationRecord.Versions.Status.FieldList[0]["documentation_versions_id"] = SQLGetInt(rdr, "documentation_versions_id");
                        }
                    }
                    catch (Exception ex)
                    {
                        BroadcastWarning("ERROR trying to create new Version record", ex);
                        return;
                    }
                }

                // Now update the record with everything from the Document Version editor (and then some)
                string strExtension = System.IO.Path.GetExtension(TxtDocumentURL.Text);  // This extracts the extension along with the "." preceding it.
                strExtension = strExtension.Right(strExtension.Length - 1);  // strip off the dot.  We don't save it in the record.

                SqlParameter[] VersionParams = new SqlParameter[10];
                VersionParams[0] = new SqlParameter("@pvintdocumentationversionsid", documentationRecord.Versions.Status.FieldList[0]["documentation_versions_id"]);
                VersionParams[1] = new SqlParameter("@pvintdocumentationid", documentationRecord.Status.FieldList[0]["documentation_id"]);
                VersionParams[2] = new SqlParameter("@pvbinembeddedfile", LoadEmbeddedFile(TxtDocumentURL.Text));
                VersionParams[3] = new SqlParameter("@pvchrfileextension", strExtension);
                VersionParams[4] = new SqlParameter("@pvchrdescription", TxtDocumentationVersionDescription.Text);
                VersionParams[5] = new SqlParameter("@pvchrversionby", UserInfo.Username);  // redundant with modfied_by field
                VersionParams[6] = new SqlParameter("@pvdatversiondatetime", DTDocumentationVersionDate.Text);
                VersionParams[7] = new SqlParameter("@pvbitactive_flag", (RadDocumentationVersionActive.Checked ? 1 : 0));
                VersionParams[8] = new SqlParameter("@pvchrlast_modified_by", UserInfo.Username);
                VersionParams[9] = new SqlParameter("@pvdatmodifieddate", DateTime.Now);
                SQLQuery("Proc_Update_Documentation_Versions", VersionParams);

                // Refresh the grid
                ClearDocumentVersionEditPanel();
                PopulateDocumentationVersionsGrid(documentationRecord);

                SetButtonState(CmdDocumentVersionSave, false, Color.Black, Color.White);
                documentationRecord.Versions.Status.IsSaved = true;
                documentationRecord.Versions.Status.IsNew = false;

                // De-render the edit panels (Doc panel gets hidden only if the doc panel's SAVE button invoked this)
                PnlDocumentation.Visible = parentGenerated;
                PnlDocumentationVersion.Visible = false;

            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to update Version record", ex);
            }
        }

        private byte[] LoadEmbeddedFile(string pathname)
        {
            if (pathname.Trim().Length <= 0)
            {
                return (new byte[1] { 0 });
            }

            try
            {
                byte[] bytEmbeddedFile = System.IO.File.ReadAllBytes(pathname);
                return (bytEmbeddedFile);
            }
            catch (Exception ex)
            {
                BroadcastWarning("Unable to read file " + pathname, ex);
                return (new byte[1] { 0 });
            }
        }

#endregion

#region ----Documentation Events

        private void GrdDocumentation_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            try
            {
                // Get the row index for whatever row was just clicked.  
                int row = e.RowIndex;
                if (row < 0)
                {
                    // header was clicked (for sorting).  Toggle the sort order direction
                    DocumentationGridSortColumn = e.ColumnIndex;
                    DocumentationGridSortOrder = (GrdDocumentation.Columns[DocumentationGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                    return;
                }

                // If a new document has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (DocumentationRecord != null)
                {
                    if (!DocumentationRecord.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE DOCUMENTATION RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveDocumentRecord(DocumentationRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the edit template. 
                DocumentationRecord = new DocumentationClass();
                PnlDocumentation.Visible = true;
                DisplayDocumentationRecord(DocumentationRecord, (int)GrdDocumentation.Rows[row].Cells["documentation_id"].Value);
                DocumentationRecord.Status.IsNew = false;
                DocumentationRecord.Versions.Status.IsSaved = true; // Indicates that no edits to the version record have occurred yet
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR while clicking an entry on the document grid", ex);
            }
        }

        private void GrdDocumentationVersion_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            try
            {
                // Get the row index for whatever row was just clicked.  
                int row = e.RowIndex;
                if (row < 0)
                {
                    // header was clicked (for sorting).  Toggle the sort order direction
                    DocumentationVersionsGridSortColumn = e.ColumnIndex;
                    DocumentationVersionsGridSortOrder = (GrdDocumentationVersion.Columns[DocumentationVersionsGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                    return;
                }

                // If a new version has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (DocumentationRecord != null)
                {
                    if (!DocumentationRecord.Versions.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current version record in the edit window has NOT been saved; do you want to save it first?", "SAVE VERSION RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveDocumentVersionRecord(DocumentationRecord, false); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // Load the selected grid data into the version edit template. 
                DocumentationRecord.Versions = new DocumentationVersionClass();
                PnlDocumentationVersion.Visible = true;
                DisplayDocumentationVersionRecord(DocumentationRecord, (int)GrdDocumentationVersion.Rows[row].Cells["documentation_Versions_id"].Value);
                DocumentationRecord.Versions.Status.IsNew = false;
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR while clicking an entry on the version grid", ex);
            }
        }

        private void CmdDocumentNew_Click(object sender, EventArgs e)
        {

            try
            {
                // If a new Document record has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (DocumentationRecord != null)
                {
                    if (!DocumentationRecord.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE DOCUMENT RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveDocumentRecord(DocumentationRecord); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // enable user edit window for populating the new record
                ClearDocumentEditPanel();

                // Create a new document record
                DocumentationRecord = new DocumentationClass();

                // On new document:
                // Reveal document and version panels
                PnlDocumentation.Visible = true;
                PnlDocumentationVersion.Visible = true;
                TxtDocumentationVersionDescription.Text = "<not specified>";

                // Disable save button and version grid 
                GrdDocumentationVersion.Enabled = false;
                SetButtonState(CmdDocumentSave, false, Color.Black, Color.White);
                SetButtonState(CmdDocumentVersionSave, false, Color.Black, Color.White);

                // Then Wait for user entry on the edit panel(s).  
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to create new document", ex);
            }
        }

        private void CmdDocumentVersionNew_Click(object sender, EventArgs e)
        {
            try
            {
                // If a new Version record has been created but not saved, prompt the user before populating the 
                //   edit area from the grid.

                if (DocumentationRecord != null)
                {
                    if (!DocumentationRecord.Versions.Status.IsSaved)
                    {
                        DialogResult dlgResult = MessageBox.Show("The current version in the edit window has NOT been saved; do you want to save it first?", "SAVE VERSION RECORD?", MessageBoxButtons.YesNoCancel);
                        switch (dlgResult)
                        {
                            case DialogResult.Yes: SaveDocumentVersionRecord(DocumentationRecord, false); break;
                            case DialogResult.No: break;
                            case DialogResult.Cancel: return;
                        }
                    }
                }

                // enable user edit window for populating the new record
                ClearDocumentVersionEditPanel();

                // Create a new version record
                DocumentationRecord.Versions = new DocumentationVersionClass();

                // Render the edit panel
                PnlDocumentationVersion.Visible = true;
                TxtDocumentationVersionDescription.Text = "<not specified>";
                SetButtonState(CmdDocumentVersionSave, false, Color.Black, Color.White);

                // Then Wait for user entry on the edit panel(s).  
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying to create a new version", ex);
            }
        }

        private void CmdDocumentationVersionBrowse_Click(object sender, EventArgs e)
        {
            // Get the URL of the file to be embedded within this documentation record
            try
            {
                OpenFileDialog dlg = new OpenFileDialog
                {
                    Title = "Retrieve Document",
                    InitialDirectory = @"C:\",
                    Filter = "Documents (*.doc,*.docx,*.pdf,*.txt)|*.doc;*.docx;*.pdf;*.txt|Spreadsheets (*.xls,*.xlsx)|*.xls;*.xlsx|All Files (*.*)|*.*",
                    RestoreDirectory = true
                };
                DialogResult = dlg.ShowDialog();
                if (DialogResult == DialogResult.OK)
                {
                    TxtDocumentURL.Text = dlg.FileName;
                    CmdDocumentationViewDocument.Enabled = true;
                }
                else
                {
                    TxtDocumentURL.Text = "";
                    CmdDocumentationViewDocument.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR trying to browse for file", ex);
            }
        }

        private void CmdDocumentationViewDocument_Click(object sender, EventArgs e)
        {
            try
            {
                if (System.IO.File.Exists(TxtDocumentURL.Text))
                {
                    System.Diagnostics.Process.Start(TxtDocumentURL.Text);
                }
                else
                {
                    MessageBox.Show("Could not find file " + TxtDocumentURL.Text);
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("Error trying to view document " + TxtDocumentURL.Text, ex);
            }
        }

        private void CmdDocumentationSearch_Click(object sender, EventArgs e)
        {
            // Search the Documentation URL and Subject fields for anything LIKE what was entered in the text box.
            //   NOTE that this was not created as a stored procedure becasue the current design of the (very popular) PopulateGrid function
            //   won't support queries like this.

            string SelectStr = "";
            try
            {
                string SearchStr = "%" + TxtDocumentationSearch.Text + "%";
                SelectStr = "SELECT * FROM DOCUMENTATION WHERE (subject LIKE '" + SearchStr + "') OR (url LIKE '" + SearchStr + "')";

                PopulateGrid(GrdDocumentation, SelectStr, CommandType.Text, RadDocumentationVersionFilterActive);
                if (GrdDocumentation.Rows.Count == 0)
                {
                    MessageBox.Show("No records matching this pattern were found");
                }
            }
            catch (Exception ex)
            {
                BroadcastWarning("ERROR:  Unable to successfully use query " + SelectStr + " during document search", ex);
            }
        }

        private void TxtDocumentationSearch_TextChanged(object sender, EventArgs e)
        {
            // Enable the documentation search button only if this text box is non-blank.
            //  (Search includes a search of the Subject and URL fields, and will cause the document grid to be filtered
            //    to those documents containing the search string)
            CmdDocumentationSearch.Enabled = (TxtDocumentationSearch.Text.Trim().Length > 0);
        }

        private void CmdViewDocumentVersion_Click(object sender, EventArgs e)
        {
            // Extract the current version from the selected version record and put it up.
            // This is a two-step process:
            //    Extract the file and save it as a temporary.  The only document name we have available to us is
            //       the main document record's URL.  If this field is empty we'll save the file under the name "NoName"
            //    Open the file for viewing.  It will open as long as its extension is associated with an existing application.

            // NOTE that because we do NOT wait for the file to be closed prior to returning control to this application,
            //    we cannot delete the saved file (since it is in use until the user closes it).
            //    It will therefore remain on disk.  The location is set to the MyDocuments folder.
            // RATIONALE FOR THIS:  User might want to compare 2 different file versions.  The asynchronous file viewing process 
            //    will allow this to occur.

            string strTempFullName = "";
            try
            {
                SqlParameter[] VersionParams = new SqlParameter[1];
                VersionParams[0] = new SqlParameter("@pvintDocumentationVersionsID", DocumentationRecord.Versions.Status.FieldList[0]["documentation_versions_id"]);
                using (SqlDataReader rdr = SQLQuery("Proc_Select_Documentation_Versions", VersionParams))
                {
                    // There should be exactly one record extracted (since documentation_versions_id is a unique identifier)
                    if (rdr.HasRows)
                    {
                        rdr.Read();
                        // Write this file to a temporary location (MyDocuments))
                        string strTempFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\";
                        string strTempFileName = ((TxtDocumentURL.Text.Trim().Length > 0) ? System.IO.Path.GetFileNameWithoutExtension(TxtDocumentURL.Text) : "NoName");
                        string strTempFileExtension = "." + rdr["file_extension"].ToString();
                        strTempFullName = strTempFilePath + strTempFileName + " " + rdr["documentation_versions_id"].ToString() + strTempFileExtension; // Version ID makes this a unique doc

                        if (System.IO.File.Exists(strTempFullName))
                        {
                            System.IO.File.Delete(strTempFullName);
                        }

                        byte[] byteArray = (byte[])rdr["embedded_file"];
                        System.IO.FileStream f = new System.IO.FileStream(strTempFullName, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write);
                        f.Write(byteArray, 0, byteArray.Length);
                        f.Flush();
                        f.Close();
                    }
                    else
                    {
                        MessageBox.Show("This version has no valid embedded document");
                        return;
                    }
                }

                // File has been saved in MyDocuments.  Open it for viewing

                if (System.IO.File.Exists(strTempFullName))
                {
                    System.Diagnostics.Process p = System.Diagnostics.Process.Start(strTempFullName);
                    //p.WaitForInputIdle();  // wait until process finishes loading
                    //p.WaitForExit(); // TBD enable this if we want to freeze until file is closed.
                }
                else
                {
                    BroadcastError("Could not find newly-extracted file " + strTempFullName, null);
                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR:  Failure trying to open file " + strTempFullName, ex);
            }
        }

        private void CmdDocumentSave_Click(object sender, EventArgs e)
        {
            // Save the document
            SaveDocumentRecord(DocumentationRecord);
        }

        private void CmdDocumentVersionSave_Click(object sender, EventArgs e)
        {
            // Save the document version
            SaveDocumentVersionRecord(DocumentationRecord, false);
        }

        private void RadDocumentationActiveStatus_Changed(object sender, EventArgs e)
        {
            // Populate the document grid
            PopulateDocumentationGrid();
        }

        private void RadDocumentationVersionActiveStatus_Changed(object sender, EventArgs e)
        {
            // Populate the document's version grid (with just the selected document's versions)
            PopulateDocumentationVersionsGrid(DocumentationRecord);
        }

        private void Documentation_TextChanged(object sender, EventArgs e)
        {
            // On any documentation text change, render the version grid and enable the save button
            GrdDocumentationVersion.Enabled = true;
            MarkAsDirty(sender, DocumentationRecord.Status, CmdDocumentSave);
        }

        private void Documentation_RadioButtonChanged(object sender, EventArgs e)
        {
            // Enable the document SAVE button
            MarkAsDirty(sender, DocumentationRecord.Status, CmdDocumentSave);
        }

        private void DocumentationVersions_TextChanged(object sender, EventArgs e)
        {
            // Enable the version SAVE button
            MarkAsDirty(sender, DocumentationRecord.Versions.Status, CmdDocumentVersionSave);
        }

        private void DocumentationVersions_RadioButtonChanged(object sender, EventArgs e)
        {
            // Enable the version SAVE button
            MarkAsDirty(sender, DocumentationRecord.Versions.Status, CmdDocumentVersionSave);
        }

#endregion

#endregion

    }
}


