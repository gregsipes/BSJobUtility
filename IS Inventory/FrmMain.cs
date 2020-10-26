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

namespace IS_Inventory
{
    public partial class FrmMain : Form
    {
        #region General

        #region General Declarations
        const string JobName = "IS Inventory";
        ActiveDirectory UserInfo = new ActiveDirectory();
        VersionStatusBar StatusBar;
        bool IsInitializing;

        #endregion

        #region General Initialization

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

            // The SubtabDataEntry_Layout event will occur immediately after we leave this function.
        }

        private void SubtabDataEntry_Layout(object sender, LayoutEventArgs e)
        {
            // This is only needed to initialize the first (visible) tab.
            //  However, this event is invoked a large number of times during loading/initialization.  Luckily this process is very slim.
            if (IsInitializing)
            {
                SubtabDataEntry.SelectedIndex = 0;
                SubtabDataEntry_SelectedIndexChanged(sender, e);
            }

            // This is the last event during initialization.  Mark initialization as complete.
            IsInitializing = false;
        }

        #endregion

        #region Safe Value Assignments
        private void SafeTextBox(TextBox t, string s)
        {
            try
            {
                t.Text = s;
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate textbox from string:  " + ex.ToString(), JobName);
                t.Text = "";
            }
        }

        private void SafeTextBox(ComboBox t, string s)
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

        private string SafeTextBoxStr(SqlDataReader rdr, string s)
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

        private int SafeTextBoxInt(SqlDataReader rdr, string s)
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

        #region SQL

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
                        // Check other SQL types like DATETIME and BIT!!!!
                        throw new NotImplementedException();
                }
            }
            else
            {
                return (null);
            }
        }

        #endregion

        #region General Helper Functions
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
                        // This record is alread in the selected table so just return true (it's already a member)
                        returnstatus = true;
                    }
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR adding value to list table " + tableName + ", field " + fieldName + ": " + ex.ToString(), JobName);
                MessageBox.Show("ERROR adding value to list table " + tableName + ", field " + fieldName + ": " + ex.ToString());
            }
            return (returnstatus);
        }

        /// <summary>
        /// Clears a combo box without deleting its list items.
        /// </summary>
        /// <param name="cmb"></param>
        private void ClearCombo(ComboBoxUnlocked cmb)
        {
            // This clears a combo box without deleting its list items.
            cmb.Text = String.Empty;
            cmb.SelectedIndex = -1;
            cmb.SelectedValue = -1;
        }

        private void ClearCombo(ComboBox cmb)
        {
            // This clears a combo box without deleting its list items.
            cmb.Text = String.Empty;
            cmb.SelectedIndex = -1;
            cmb.SelectedValue = -1;
        }


        private void SubtabDataEntry_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (SubtabDataEntry.SelectedIndex)
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
                        PopulateServersGrid();
                    }
                    break;
                case 3:  // Documentation
                    if (!DocumentationInitialized)
                    {
                        DocumentationInitialized = true;
                        PopulateDocumentationGrid();
                    }
                    break;
                default:
                    break;

            }
        }

        
        #endregion

        #region General Events
        private void FrmMain_Load(object sender, EventArgs e)
        {
            // TBD - Nothing yet needed here so it's just a placeholder for now
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

        #region HW Tab Declarations
        HardwareTemplateClass HWTemplate;
        bool HardwareTemplateTabInitialized = false;

        private void InitializeHardwareTemplate()
        {
            PnlHardwareTemplate.Visible = false;
        }

        #endregion

        #region Populate HW Tab Templates
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

                PopulateGrid(GrdHardwareTemplate, "Proc_Select_Hardware_Template", CommandType.StoredProcedure);
                HardwareTemplateTabInitialized = true;
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate Hardware Template tab: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate Hardware Template tab: " + ex.ToString());
            }
            return;
        }
        #endregion

        #region Populate HW Tab Combo Boxes and Grids
        /// <summary>
        /// Populates the selected combobox.
        /// </summary>
        /// <param name="cmb">Name of the combobox</param>
        /// <param name="tableName">Name of the table whose data will be loaded into the combobox</param>
        /// <param name="displayMember">Name of the field in the table to be displayed</param>
        /// <param name="valueMember">Name of the primary key field (must be a unique ID)</param>
        private void PopulateComboBox(ComboBox cmb, string tableName, string displayMember, string valueMember)
        {
            // This function constructs the following query:
            //    SELECT * FROM <tablename> ORDER BY <display member> (for all active records)
            // and uses it as the data source for the specified combo box.  
            // displayMember must be a valid field within the dataset, and it is the field that will be displayed in the combo box.

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
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate combobox " + cmb.Name + ": " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate combobox " + cmb.Name + ": " + ex.ToString());
            }
        }

        /// <summary>
        /// Populates the selected combobox.
        /// </summary>
        /// <param name="cmb">Name of the combobox</param>
        /// <param name="tableName">Name of the table whose data will be loaded into the combobox</param>
        /// <param name="displayMember">Name of the field in the table to be displayed</param>
        /// <param name="valueMember">Name of the primary key field (must be a unique ID)</param>
        private void PopulateComboBox(ComboBoxUnlocked cmb, string tableName, string displayMember, string valueMember)
        {
            // This function constructs the following query:
            //    SELECT * FROM <tablename> ORDER BY <display member> (for all active records)
            // and uses it as the data source for the specified combo box.  
            // displayMember must be a valid field within the dataset, and it is the field that will be displayed in the combo box.

            try
            {
                PopulateComboBox((ComboBox)cmb, tableName, displayMember, valueMember);
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate unlocked combobox " + cmb.Name + ": " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate unlocked combobox " + cmb.Name + ": " + ex.ToString());
            }
        }

        /// <summary>
        /// Displays query results on a generic grid 
        /// (If clearDataSource is true, lets Windows decide column sizes and row heights, 
        ///   and displays all fields in the query).
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="procName"></param>
        void PopulateGrid(DataGridView dgv, string procName, CommandType command)
        {
            try
            {
                // Populate the specified grid using a generic approach - unilaterally fill it left to right
                //   based on the order specified in the stored procedure.

                using (SqlDataReader rdr = SQLQuery(procName, command))
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
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate grid " + dgv.Name + " with sproc " + procName + ": " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate grid " + dgv.Name + " with sproc " + procName + ": " + ex.ToString());
            }
        }

        #endregion

        #region HW Tab Data Display Functions
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
                hWTemplate.HWRecord = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Hardware_Template", TemplateParams);
                Dictionary<string, object> hwr = HWTemplate.HWRecord[0];
                SafeTextBox(TxtPCtemplate_name, hwr["template_name"].ToString());
                SafeTextBox(CmbPCcabinets_id, hwr["Cabinet"].ToString());
                SafeTextBox(CmbPCcddevices_id, hwr["CD_Device"].ToString());
                SafeTextBox(CmbPCharddrive1_id, hwr["Drive1"].ToString());
                SafeTextBox(CmbPCharddrive2_id, hwr["Drive2"].ToString());
                SafeTextBox(CmbPCkeyboards_id, hwr["Keyboard"].ToString());
                SafeTextBox(CmbPCmanufacturers_id, hwr["Manufacturer"].ToString());
                SafeTextBox(CmbPCmiscdrives_id, hwr["MiscDrive"].ToString());
                SafeTextBox(CmbPCmiscellaneouscard_id, hwr["Miscellaneous"].ToString());
                SafeTextBox(CmbPCmodels_id, hwr["Model"].ToString());
                SafeTextBox(CmbPCmonitor1_id, hwr["Monitor1"].ToString());
                SafeTextBox(CmbPCmonitor2_id, hwr["Monitor2"].ToString());
                SafeTextBox(CmbPCmotherboards_id, hwr["Motherboard"].ToString());
                SafeTextBox(CmbPCmice_id, hwr["Mouse"].ToString());
                SafeTextBox(CmbPCnics_id, hwr["NIC"].ToString());
                SafeTextBox(CmbPCprocessors_id, hwr["Processor"].ToString());
                SafeTextBox(CmbPCram_id, hwr["RAM"].ToString());
                SafeTextBox(CmbPCsoundcards_id, hwr["Sound_Card"].ToString());
                SafeTextBox(CmbPCspeakers_id, hwr["Speakers"].ToString());
                SafeTextBox(CmbPCvideocards_id, hwr["Video_Card"].ToString());
                if ((bool)hwr["active_flag"])
                {
                    RadPCactive_flag.Checked = true;
                }
                else
                {
                    RadPCinactive_flag.Checked = true;
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate hardware template editor: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate hardware template editor: " + ex.ToString());
            }
        }
        #endregion

        #region HW Tab Classes

        public class HardwareTemplateClass
        {
            // A simple class containing a dictionary containing all fields in the TempatesHardware table.
            public List<Dictionary<string, object>> HWRecord; // Conforms to DataIO dictionary return type.
            public bool IsActive;
            public bool IsSaved;

            public HardwareTemplateClass()
            {
                HWRecord = new List<Dictionary<string, object>>();
                IsSaved = true;
                IsActive = true;

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> hwrecord = new Dictionary<string, object>
                {
                    ["hardware_templates_ID"] = 0,
                    ["template_name"] = ""
                };
                HWRecord.Add(hwrecord);
            }
        }

        #endregion

        #region HW Tab Helper Functions

        /// <summary>
        /// Executed when the user hits the ENTER key or when focus moves off the control.  
        /// </summary>
        /// <param name="cmb"></param>
        /// <returns></returns>
        private bool ConfirmComboboxEntry(ComboBox cmb)
        {
            // Invoked when the user hits the ENTER key OR when focus moves off the control.
            //   When this happens we want to check if the user entered a text value not already on the combobox list.
            //   Anything not already on the list deserves a prompt asking if it should be added to the list.
            //   Once it's added we need to reload the combo box.
            // NOTE:  Combobox items list is not IEnumerable to LINQ doesn't work directly on this list.
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
                    case "cmbpcharddrive1_id": recordadded = AddToTableIfNotAMember(cmb, "lstDrives", "Drives", "Drives_ID", newdata); break;
                    case "cmbpcharddrive2_id": recordadded = AddToTableIfNotAMember(cmb, "lstDrives", "Drives", "Drives_ID", newdata); break;
                    case "cmbpckeyboards_id": recordadded = AddToTableIfNotAMember(cmb, "lstKeyboards", "Keyboard", "Keyboards_ID", newdata); break;
                    case "cmbpcmanufacturers_id": recordadded = AddToTableIfNotAMember(cmb, "lstManufacturers", "Manufacturer", "Manufacturers_ID", newdata); break;
                    case "cmbpcmiscdrives_id": recordadded = AddToTableIfNotAMember(cmb, "lstDrives", "Drives", "Drives_ID", newdata); break;
                    case "cmbpcmiscellaneouscard_id": recordadded = AddToTableIfNotAMember(cmb, "lstMiscellaneous", "Miscellaneous", "Miscellaneous_ID", newdata); break;
                    case "cmbpcmodels_id": recordadded = AddToTableIfNotAMember(cmb, "lstModels", "Model", "Models_ID", newdata); break;
                    case "cmbpcmonitor1_id": recordadded = AddToTableIfNotAMember(cmb, "lstMonitors", "Monitor", "Monitors_ID", newdata); break;
                    case "cmbpcmonitor2_id": recordadded = AddToTableIfNotAMember(cmb, "lstMonitors", "Monitor", "Monitors_ID", newdata); break;
                    case "cmbpcmotherboards_id": recordadded = AddToTableIfNotAMember(cmb, "lstMotherboards", "Motherboard", "Motherboards_ID", newdata); break;
                    case "cmbpcmice_id": recordadded = AddToTableIfNotAMember(cmb, "lstMice", "Mouse", "Mice_ID", newdata); break;
                    case "cmbpcnics_id": recordadded = AddToTableIfNotAMember(cmb, "lstNICs", "NIC", "NICs_ID", newdata); break;
                    case "cmbpcprocessors_id": recordadded = AddToTableIfNotAMember(cmb, "lstProcessors", "Processor", "Processors_ID", newdata); break;
                    case "cmbpcram_id": recordadded = AddToTableIfNotAMember(cmb, "lstRAM", "RAM", "RAM_ID", newdata); break;
                    case "cmbpcsoundcards_id": recordadded = AddToTableIfNotAMember(cmb, "lstSoundCards", "Sound_Card", "Sound_Cards_ID", newdata); break;
                    case "cmbpcspeakers_id": recordadded = AddToTableIfNotAMember(cmb, "lstSpeakers", "Speakers", "Speakers_ID", newdata); break;
                    case "cmbpcvideocards_id": recordadded = AddToTableIfNotAMember(cmb, "lstVideoCards", "Video_Card", "Video_Cards_ID", newdata); break;
                    default:
                        DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to process combobox " + cmb.Name + " with data " + newdata, JobName);
                        MessageBox.Show("ERROR trying to process combobox " + cmb.Name + " with data " + newdata);
                        break;
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR processing combobox " + cmb.Name + ", data = " + newdata + ": " + ex.ToString(), JobName);
                MessageBox.Show("ERROR processing combobox " + cmb.Name + ", data = " + newdata + ": " + ex.ToString());
            }
            return (recordadded);
        }

        private void SaveHardwareTemplate(HardwareTemplateClass hWTemplate)
        {

            // I thought abour redoing this using the HWTemplate's dictionary.  Put everything in a loop.
            //  Unfortunately, there's much more in the dictionary than just the fields in the table,
            //  so that might be a really tough thing to work out.

            SqlParameter[] HWParams = new SqlParameter[23];
            HWParams[0] = new SqlParameter("@pvintTemplateID", hWTemplate.HWRecord[0]["hardware_templates_id"]);
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

            HWTemplate.IsSaved = true;
            SQLQuery("Proc_Update_Hardware_Template", HWParams);

            // Refresh the grid
            ClearHWTemplate();
            PopulateGrid(GrdHardwareTemplate, "Proc_Select_Hardware_Template", CommandType.StoredProcedure);
            PnlHardwareTemplate.Visible = false;
        }

        private void CmdNewHardwareTemplate_Click(object sender, EventArgs e)
        {
            // Check if there is a current record and if so, whether or not is has been saved.
            //   If not saved, the prompt to save it.

            if (HWTemplate != null)
            {
                if (!HWTemplate.IsSaved)
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
            int RecID = SafeTextBoxInt(rdr, "hardware_templates_id");
            TxtPCtemplate_name.Text = templatename + RecID.ToString();
            HWTemplate = new HardwareTemplateClass();
            HWTemplate.HWRecord[0]["hardware_templates_id"] = RecID;
            HWTemplate.HWRecord[0]["template_name"] = TxtPCtemplate_name.Text;

            // Update the new record with the (temporary) template name
            HWParams = new SqlParameter[2];
            HWParams[0] = new SqlParameter("@pvintTemplateID", RecID);
            HWParams[1] = new SqlParameter("@pvchrTemplateName", TxtPCtemplate_name.Text);
            SQLQuery("Proc_Update_Hardware_Template", HWParams);

            PnlHardwareTemplate.Visible = true;
        }

        private void ClearHWTemplate()
        {
            // Remove all selected information from each of the combo boxes in the HW Template edit panel.
            //   (Start with a clean slate)
            TxtPCtemplate_name.Text = "";
            TxtPCcomments.Text = "";

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

        #region HW Tab Events

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

            if (HWTemplate != null)
            {
                if (!HWTemplate.IsSaved)
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


        /// <summary>
        /// An event unique to ComboBoxUnlocked.  It fires when the user depresses the ENTER key or when the combobox loses focus.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Event_DataEntryComplete(object sender, EventArgs e)
        {
            ComboBox cmb = (ComboBox)sender;

            // If this entry isn't on the list then add it
            bool recordadded = ConfirmComboboxEntry(cmb);

            // If the combobox entry was accepted (added to the table or already a member of the table) then update its dictionary entry in the template record
            // (fieldname is identical to combobox name EXCEPT for the leading 5 letters (in this case, "CMBPC" needs to be stripped)
            if (recordadded)
            {
                string keyname = cmb.Name.Right(cmb.Name.Length - 5);
                HWTemplate.HWRecord[0][keyname] = cmb.SelectedValue;
            }
        }

        /// <summary>
        /// Save the current HW template edit panel to the TemplatesHardware table
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmbSaveHardwareTemplate_Click(object sender, EventArgs e)
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

        private void RadPCactive_flag_CheckedChanged(object sender, EventArgs e)
        {
            // Mark whether or not the record is rendered active.  Inactive records won't appear in other lists.
            HWTemplate.IsActive = RadPCactive_flag.Checked;
        }
        #endregion

        #endregion

        #region Category Edit

        #region Category Edit Declarations
        CategoryEditClass CategoryEdit;

        private void InitializeCategoryEdit()
        {
            PnlCategoryEditNewSave.Visible = false;
            PnlCategoryEditItem.Visible = false;
            CmdCategoryEditSaveItem.Enabled = false;
        }
        #endregion

        #region Category Edit Classes

        public class CategoryEditClass
        {
            // A simple class containing a dictionary containing all fields in the lst<xxx> categories tables.
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

        #region Category Edit Data Display Functions
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
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate Category editor: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate Category editor: " + ex.ToString());
            }
        }

        private void ClearCategoryEditTemplate()
        {
            // Remove all selected information from the Category edit panel.
            //   (Start with a clean slate)
            TxtItemName.Text = "";
            RadCategoryActive.Checked = true;
        }

        #endregion

        #region Category Edit Events
        private void CmbSelectedList_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Populate the grid with whatever selection was made here
            SelectCategoryListToPopulate();

            //Render the New/Save panel visible
            PnlCategoryEditNewSave.Visible = true;
        }

        private void SelectCategoryListToPopulate()
        {
            string listname = CmbSelectedList.Text; // The names in this combobox list MUST be exactly the same as the tablenames, less the "lst" prefix.
            string SelectStr = "SELECT * FROM lst" + listname + " ";
            if (RadCategoryActiveOnly.Checked)
            {
                SelectStr += " WHERE active_flag = 1";
            }
            if (RadCategoryInactiveOnly.Checked)
            {
                SelectStr += " WHERE active_flag = 0";
            }

            PopulateGrid(GrdSelectedList, SelectStr, CommandType.Text);
            FormatCategoryGrid(GrdSelectedList);
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
            // Check if a previous record needs to be saved before creating a new record!
            CheckForEditedCategoryItem(CategoryEdit);

            // Create a new category record and initialize the editor
            CategoryEdit = new CategoryEditClass(GrdSelectedList, CmbSelectedList.Text);
            FormatCategoryGrid(GrdSelectedList);
            ClearCategoryEditTemplate();
            PnlCategoryEditItem.Visible = true;
        }

        private void GrdSelectedList_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;

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
            CmdCategoryEditSaveItem.Enabled = false;
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
                        case DialogResult.Yes:
                            SaveCategoryEdit(CategoryEdit);
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
            CmdCategoryEditSaveItem.Enabled = false;
            categoryEdit.IsSaved = true;
            categoryEdit.IsNew = false;
        }

        private void RadCategoryActiveOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (CmbSelectedList.Text.Length > 0)
            {
                CmbSelectedList_SelectedIndexChanged(sender, e);
            }
        }

        private void RadCategoryInactiveOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (CmbSelectedList.Text.Length > 0)
            {
                CmbSelectedList_SelectedIndexChanged(sender, e);
            }
        }

        private void RadCategoryBoth_CheckedChanged(object sender, EventArgs e)
        {
            if (CmbSelectedList.Text.Length > 0)
            {
                CmbSelectedList_SelectedIndexChanged(sender, e);
            }
        }

        private void TxtItemName_TextChanged(object sender, EventArgs e)
        {
            // This text box is limited to 100 characters - same length as maximum allowed in the category list tables
            CmdCategoryEditSaveItem.Enabled = true;
            CategoryEdit.IsSaved = false;
        }

        private void RadCategoryActive_CheckedChanged(object sender, EventArgs e)
        {
            CmdCategoryEditSaveItem.Enabled = true;
            CategoryEdit.IsSaved = false;
        }

        #endregion

        #endregion

        #region IP Address Edit

        #region IP Address Edit Declarations

        IPAddressEditClass IPAddressRecord;
        ListSortDirection IPAddressGridSortOrder;
        int IPAddressGridSortColumn;
        bool IPAddressesInitialized;

        private void InitializeIPAddressesEdit()
        {
            // By default the SAVE button and the IP address text box are not enabled.
            //   (Existing IP addresses are not editable, and SAVE is enabled only after an edit)

            CmdSaveIPAddress.Enabled = false;
            TxtIPAddress.Enabled = false;
            PnlIPEditItem.Visible = false;
            IPAddressGridSortOrder = ListSortDirection.Ascending;
            IPAddressGridSortColumn = 1; // Initially sorted by IP Address
            IPAddressesInitialized = false;
        }

        #endregion

        #region IP Address Edit Classes

        public class IPAddressEditClass
        {
            // A simple class containing all fields in the IPAddresses table
            public bool IsSaved;
            public bool IsNew;
            public List<Dictionary<string, object>> FieldList; // Conforms to DataIO dictionary return type.

            public IPAddressEditClass()
            {
                FieldList = new List<Dictionary<string, object>>();
                IsSaved = true;
                IsNew = true;

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
                FieldList.Add(fields);
            }
        }
        #endregion

        #region IP Address Edit Data Display Functions

        private void PopulateIPAddressGrid()
        {
            PopulateGrid(GrdIPAddresses, "Proc_Select_IP_Addresses", CommandType.StoredProcedure);
            CmdSaveIPAddress.Enabled = false;

            // For this grid we need to specify column attributes
            GrdIPAddresses.Columns[0].Visible = false;   // Hide the IP Address ID
            GrdIPAddresses.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;  // The NOTES column will be fixed length
            GrdIPAddresses.Columns[5].Width = 500;
            GrdIPAddresses.Sort(GrdIPAddresses.Columns[IPAddressGridSortColumn], IPAddressGridSortOrder);
        }

        private void DisplayIPAddressRecord(IPAddressEditClass iPAddressRecord, int recordID)
        {
            // Get all data associated with this record
            try
            {
                ClearIPAddressEditPanel();
                SqlParameter[] IPAddressParams = new SqlParameter[1];
                IPAddressParams[0] = new SqlParameter("@pvintRecordID", recordID);
                iPAddressRecord.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_IP_Addresses", IPAddressParams);
                Dictionary<string, object> ipr = iPAddressRecord.FieldList[0];
                SafeTextBox(TxtIPAddress, ipr["IPAddress"].ToString());
                SafeTextBox(TxtIPVLAN, ipr["VLAN"].ToString());
                SafeTextBox(TxtIPDescription, ipr["Description"].ToString());
                SafeTextBox(TxtIPTranslatedAddress, ipr["TranslatedAddr"].ToString());
                SafeTextBox(TxtIPNotes, ipr["Notes"].ToString());
                if ((bool)ipr["IsActive"])
                {
                    RadIPActive.Checked = true;
                }
                else
                {
                    RadIPInactive.Checked = true;
                }

                // Disable the SAVE button until the user makes an edit.
                CmdSaveIPAddress.Enabled = false;
                iPAddressRecord.IsSaved = true;

                // Existing IP addresses are NOT editable, so disable the IPAddress text box
                TxtIPAddress.Enabled = false;
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to populate IP address record editor: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to populate IP address record editor: " + ex.ToString());
            }
        }

        private void ClearIPAddressEditPanel()
        {
            // Remove all selected information from each of the text boxes in the IP Address edit panel.
            //   (Start with a clean slate)
            TxtIPAddress.Text = "";
            TxtIPDescription.Text = "";
            TxtIPNotes.Text = "";
            TxtIPTranslatedAddress.Text = "";
            TxtIPVLAN.Text = "";
            RadIPActive.Checked = true;
        }

        #endregion

        #region IP Address Edit Helper Functions

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
                                        iP.FieldList[0]["ipaddresses_id"] = SafeTextBoxInt(rdr, "ipaddresses_id");
                                        IsNew = true;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to create new IP Address record: " + ex.ToString(), JobName);
                                MessageBox.Show("ERROR trying to create new IP Address record: " + ex.ToString());
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
                    IPParams[0] = new SqlParameter("@pvintRecordID", iP.FieldList[0]["ipaddresses_id"]);
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
                    iP.IsSaved = true;
                }

            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to save new IP Address record: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to save new IP Address record: " + ex.ToString());
            }

        }

        private bool IsDuplicateIPAddress(string iPAddress)
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
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR translating IP Address to octets: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR translating IP Address to octets: " + ex.ToString());
                return ("");
            }
        }

        #endregion

        #region IP Address Events

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

            // If a new template entry has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            if (IPAddressRecord != null)
            {
                if (!IPAddressRecord.IsSaved)
                {
                    DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE IP ADDRESS RECORD?", MessageBoxButtons.YesNoCancel);
                    switch (dlgResult)
                    {
                        case DialogResult.Yes:
                            // Save record first  
                            SaveIPAddressRecord(IPAddressRecord);
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
            // Load the selected grid data into the edit template. 
            IPAddressRecord = new IPAddressEditClass();
            PnlIPEditItem.Visible = true;
            DisplayIPAddressRecord(IPAddressRecord, (int)GrdIPAddresses.Rows[row].Cells["ipaddresses_id"].Value);
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
            CmdSaveIPAddress.Enabled = true;
            IPAddressRecord.IsSaved = false;
        }

        private void CmdNewIPAddress_Click(object sender, EventArgs e)
        {
            // Enable the IPAddress text box
            TxtIPAddress.Enabled = true;

            // On New, render the IP Edit panel
            PnlIPEditItem.Visible = true;

            // If a new template entry has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            if (IPAddressRecord != null)
            {
                if (!IPAddressRecord.IsSaved)
                {
                    DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE IP ADDRESS RECORD?", MessageBoxButtons.YesNoCancel);
                    switch (dlgResult)
                    {
                        case DialogResult.Yes:
                            // Save record first  
                            SaveIPAddressRecord(IPAddressRecord);
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
        #endregion

        #endregion

        #region PCs and MACs

        #region PCs and MACs Declarations

        PCsAndMACsEditClass PCsAndMacsRecord;
        ListSortDirection PCsAndMacsGridSortOrder;
        int PCsAndMacsGridSortColumn;
        bool PCsAndMacsInitialized;

        private void InitializePCsAndMACs()
        {
            PnlPCUser.Visible = false;   // Render this after a NEW or an existing selection is made
            PnlPCHardware.Visible = false; // Render this only after a HW template has been selected
            CmdSavePC.Enabled = false;

            // Populate Department, PC Type and PC Template comboboxes
            PopulateComboBox(CmbPCTemplate, "TemplatesHardware", "template_name", "hardware_templates_id");
            PopulateComboBox(CmbPCType, "lstPCTypes", "pctype", "PCType_id");
            PopulateComboBox(CmbPCDepartment, "lstDepartments", "department", "departments_id");

        }

        #endregion

        #region PCs and MACs Classes
        public class PCsAndMACsEditClass
        {
            // A simple class containing all fields in the IPAddresses table
            public bool IsSaved;
            public bool IsNew;
            public List<Dictionary<string, object>> FieldList; // Conforms to DataIO dictionary return type.

            public PCsAndMACsEditClass()
            {
                FieldList = new List<Dictionary<string, object>>();
                IsSaved = true;
                IsNew = true;

                // On instantiation, create all necessary dictionary keys in the first record (should be the ONLY record ever created on this list) 
                Dictionary<string, object> fields = new Dictionary<string, object>
                {
                    ["pcmac_id"] = 0,
                    ["system_id"] = "",
                    ["active_flag"] = true
                };
                FieldList.Add(fields);
            }
        }

        #endregion

        #region PCs and MACs Data Display Functions

        private void DisplayPCsAndMACsRecord(PCsAndMACsEditClass pCMac, int recordID)
        {
            // Get all data associated with this record
            try
            {
                ClearPCsAndMACsEditPanel();

                SqlParameter[] PCMacParams = new SqlParameter[1];
                PCMacParams[0] = new SqlParameter("@pvintPCMacID", recordID);
                pCMac.FieldList = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_PCsAndMACs", PCMacParams);
                Dictionary<string, object> pcm = pCMac.FieldList[0];

                SafeTextBox(CmbPCType, pcm["pc_type"].ToString());
                SafeTextBox(TxtPCSystemID, pcm["system_id"].ToString());
                SafeTextBox(TxtPCUsername, pcm["username"].ToString());
                SafeTextBox(CmbPCDepartment, pcm["user_department"].ToString());

                bool templateidokay = int.TryParse(pcm["pc_template_id"].ToString(), out int templateid);
                if (templateidokay)
                {
                    SqlParameter[] HWParams = new SqlParameter[1];
                    HWParams[0] = new SqlParameter("@pvintTemplateID", templateid);
                    using (SqlDataReader rdr = SQLQuery("Proc_Select_Hardware_Template", HWParams))
                    {
                        if (rdr.HasRows)
                        {
                            rdr.Read();
                            string templatename = SafeTextBoxStr(rdr, "template_name");
                            SafeTextBox(CmbPCTemplate, templatename);
                        }
                    }
                }
                SafeTextBox(TxtPCSystemTag, pcm["system_tag"].ToString());
                SafeTextBox(TxtPCSystemSN, pcm["system_serial_number"].ToString());
                SafeTextBox(TxtPCMonitorTag, pcm["monitor_tag"].ToString());
                SafeTextBox(TxtPCMonitorSN, pcm["monitor_serial_number"].ToString());
                SafeTextBox(TxtPCKeyboardTag, pcm["keyboard_tag"].ToString());
                SafeTextBox(TxtPCKeyboardSN, pcm["keyboard_serial_number"].ToString());
                SafeTextBox(TxtPCMACAddress, pcm["macaddress"].ToString());

                SafeTextBox(TxtPCManufacturer, pcm["manufacturer"].ToString());
                SafeTextBox(TxtPCCabinet, pcm["cabinet"].ToString());
                SafeTextBox(TxtPCModel, pcm["model"].ToString());
                SafeTextBox(TxtPCProcessor, pcm["processor"].ToString());
                SafeTextBox(TxtPCRAM, pcm["ram"].ToString());
                SafeTextBox(TxtPCMotherboard, pcm["motherboards"].ToString());
                SafeTextBox(TxtPCMonitor1, pcm["monitor1"].ToString());
                SafeTextBox(TxtPCMonitor2, pcm["monitor2"].ToString());
                SafeTextBox(TxtPCHardDrive1, pcm["hard_drive1"].ToString());
                SafeTextBox(TxtPCHardDrive2, pcm["hard_drive2"].ToString());
                SafeTextBox(TxtPCMiscDrive, pcm["miscellaneous_drive"].ToString());
                SafeTextBox(TxtPCNIC, pcm["nic"].ToString());
                SafeTextBox(TxtPCCDDevice, pcm["cd_device"].ToString());
                SafeTextBox(TxtPCVideoCard, pcm["video_card"].ToString());
                SafeTextBox(TxtPCSoundCard, pcm["sound_card"].ToString());
                SafeTextBox(TxtPCMiscCard, pcm["miscellaneous_card"].ToString());
                SafeTextBox(TxtPCSpeakers, pcm["speakers"].ToString());
                SafeTextBox(TxtPCKeyboard, pcm["keyboard"].ToString());
                SafeTextBox(TxtPCMouse, pcm["mouse"].ToString());
                SafeTextBox(TxtPCComments2, pcm["comments"].ToString());
                if ((bool)pcm["active_flag"])
                {
                    RadPCActive.Checked = true;
                }
                else
                {
                    RadPCInactive.Checked = true;
                }

                // Disable the SAVE button until the user makes an edit.
                CmdSavePC.Enabled = false;
                pCMac.IsSaved = true; // Indicates that no edits to this record have occurred (yet)
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR displaying PC / Mac records: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR displaying PC / Mac records: " + ex.ToString());
            }
        }

        private void PopulatePCsAndMacsGrid()
        {
            PopulateGrid(GrdPCsAndMacs, "Proc_Select_PCsAndMACs", CommandType.StoredProcedure);
            CmdSavePC.Enabled = false;
            // For this grid we need to specify column attributes
            GrdPCsAndMacs.Columns[0].Visible = false;   // Hide the IP Address ID
            GrdPCsAndMacs.Sort(GrdPCsAndMacs.Columns[PCsAndMacsGridSortColumn], PCsAndMacsGridSortOrder);
        }

        private void ClearPCsAndMACsEditPanel()
        {
            // Remove all selected information from each of the text boxes in PC / MAC edit panel.
            //   (Start with a clean slate)
            ClearCombo(CmbPCType);
            TxtPCSystemID.Text = "";
            TxtPCUsername.Text = "";
            ClearCombo(CmbPCDepartment);
            ClearCombo(CmbPCTemplate);
            TxtPCtemplate_name.Text = "";
            TxtPCSystemTag.Text = "";
            TxtPCSystemSN.Text = "";
            TxtPCMonitorTag.Text = "";
            TxtPCMonitorSN.Text = "";
            TxtPCKeyboardTag.Text = "";
            TxtPCKeyboardSN.Text = "";
            TxtPCMACAddress.Text = "";
            TxtPCManufacturer.Text = "";
            TxtPCCabinet.Text = "";
            TxtPCModel.Text = "";
            TxtPCProcessor.Text = "";
            TxtPCRAM.Text = "";
            TxtPCMotherboard.Text = "";
            TxtPCMonitor1.Text = "";
            TxtPCMonitor2.Text = "";
            TxtPCHardDrive1.Text = "";
            TxtPCHardDrive2.Text = "";
            TxtPCMiscDrive.Text = "";
            TxtPCNIC.Text = "";
            TxtPCCDDevice.Text = "";
            TxtPCVideoCard.Text = "";
            TxtPCSoundCard.Text = "";
            TxtPCMiscCard.Text = "";
            TxtPCSpeakers.Text = "";
            TxtPCKeyboard.Text = "";
            TxtPCMouse.Text = "";
            TxtPCComments2.Text = "";
            RadPCActive.Checked = true;
        }

        #endregion

        #region PCs and MACs Helper Functions

        private void SavePCsAndMacsRecord(PCsAndMACsEditClass pCsMacs)
        {
            try
            {
                // Qualification tests:  
                //   SystemID and PC Type may not be null.
                if (QualifyPCsAndMacsValues(pCsMacs))
                {
                    // Is this a new record (i.e., is the IsNew property enabled?) If so, then perform a record insert prior to the update.

                    if (pCsMacs.IsNew)
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
                                    pCsMacs.FieldList[0]["pcmac_id"] = SafeTextBoxInt(rdr, "pcmac_id");
                                }

                            }
                            catch (Exception ex)
                            {
                                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to create new PC / MAC record: " + ex.ToString(), JobName);
                                MessageBox.Show("ERROR trying to create new PC / MAC record: " + ex.ToString());
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
                    PCParams[0] = new SqlParameter("@pvintpcmac_id", pCsMacs.FieldList[0]["pcmac_id"]);
                    PCParams[1] = new SqlParameter("@pvchrpc_type", CmbPCType.Text);
                    PCParams[2] = new SqlParameter("@pvchrsystem_id", TxtPCSystemID.Text);
                    PCParams[3] = new SqlParameter("@pvchrusername", TxtPCUsername.Text);
                    PCParams[4] = new SqlParameter("@pvchruser_department", CmbPCDepartment.Text);
                    PCParams[5] = new SqlParameter("@pvintpc_template_id", PCsAndMacsRecord.FieldList[0]["pc_template_id"]);
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
                    PCParams[34] = new SqlParameter("@pvdatmodifiedfate", DateTime.Now);
                    PCParams[35] = new SqlParameter("@pvchrmodifiedby", UserInfo.Username);

                    SQLQuery("Proc_Update_PCsAndMACs", PCParams);

                    // Refresh the grid
                    ClearPCsAndMACsEditPanel();
                    PopulatePCsAndMacsGrid();

                    CmdSavePC.Enabled = false;
                    PCsAndMacsRecord.IsSaved = true;
                    PCsAndMacsRecord.IsNew = false;

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
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to save new PC / MAC record: " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to save new PC / MAC record: " + ex.ToString());
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

        #region PCs and MACs Events

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
                        int templateid = SafeTextBoxInt(rdr, "hardware_templates_id");
                        PCsAndMacsRecord.FieldList[0]["pc_template_id"] = templateid;

                        SqlParameter[] TemplateParams = new SqlParameter[1];
                        TemplateParams[0] = new SqlParameter("@pvintTemplateID", templateid);
                        List<Dictionary<string, object>> hwtemplate = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.ISInventory, CommandType.StoredProcedure, "Proc_Select_Hardware_Template", TemplateParams);
                        Dictionary<string, object> hwr = hwtemplate[0];

                        // Render the (new) Template values!
                        SafeTextBox(TxtPCManufacturer, hwr["Manufacturer"].ToString());
                        SafeTextBox(TxtPCCabinet, hwr["Cabinet"].ToString());
                        SafeTextBox(TxtPCModel, hwr["Model"].ToString());
                        SafeTextBox(TxtPCProcessor, hwr["Processor"].ToString());
                        SafeTextBox(TxtPCRAM, hwr["RAM"].ToString());
                        SafeTextBox(TxtPCMotherboard, hwr["Motherboard"].ToString());
                        SafeTextBox(TxtPCMonitor1, hwr["Monitor1"].ToString());
                        SafeTextBox(TxtPCMonitor2, hwr["Monitor2"].ToString());
                        SafeTextBox(TxtPCHardDrive1, hwr["Drive1"].ToString());
                        SafeTextBox(TxtPCHardDrive2, hwr["Drive2"].ToString());
                        SafeTextBox(TxtPCMiscDrive, hwr["MiscDrive"].ToString());
                        SafeTextBox(TxtPCNIC, hwr["NIC"].ToString());
                        SafeTextBox(TxtPCCDDevice, hwr["CD_Device"].ToString());
                        SafeTextBox(TxtPCVideoCard, hwr["Video_Card"].ToString());
                        SafeTextBox(TxtPCSoundCard, hwr["Sound_Card"].ToString());
                        SafeTextBox(TxtPCMiscCard, hwr["Miscellaneous"].ToString());
                        SafeTextBox(TxtPCSpeakers, hwr["Speakers"].ToString());
                        SafeTextBox(TxtPCKeyboard, hwr["Keyboard"].ToString());
                        SafeTextBox(TxtPCMouse, hwr["Mouse"].ToString());
                        SafeTextBox(TxtPCComments2, hwr["comments"].ToString());  // Note LOWERCASE!  Dictionaries are case-sensitive.
                    }

                }

            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ERROR trying to load / populate hardware template " + CmbPCTemplate.Text + ": " + ex.ToString(), JobName);
                MessageBox.Show("ERROR trying to load / populate hardware template " + CmbPCTemplate.Text + ": " + ex.ToString());
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

            // If a new template entry has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            if (PCsAndMacsRecord != null)
            {
                if (!PCsAndMacsRecord.IsSaved)
                {
                    DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE PC / MAC RECORD?", MessageBoxButtons.YesNoCancel);
                    switch (dlgResult)
                    {
                        case DialogResult.Yes:
                            // Save record first  
                            SavePCsAndMacsRecord(PCsAndMacsRecord);
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

            // Load the selected grid data into the edit template. 
            PCsAndMacsRecord = new PCsAndMACsEditClass();
            PnlPCUser.Visible = true;
            PnlPCHardware.Visible = true;
            DisplayPCsAndMACsRecord(PCsAndMacsRecord, (int)GrdPCsAndMacs.Rows[row].Cells["pcmac_id"].Value);
            PCsAndMacsRecord.IsNew = false;
        }

        private void PCsAndMacs_TextChanged(object sender, EventArgs e)
        {
            // This event is associated with ALL text and radio boxes in the PC / MAC tab's edit menu.
            //   See the PROPERTIES View for each of the test and radio boxes
            // On any change, enable the Save button
            if (!(PCsAndMacsRecord == null))
            {
                CmdSavePC.Enabled = true;
                PCsAndMacsRecord.IsSaved = false;
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

            if (PCsAndMacsRecord != null)
            {
                if (!PCsAndMacsRecord.IsSaved)
                {
                    DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE PC / MAC RECORD?", MessageBoxButtons.YesNoCancel);
                    switch (dlgResult)
                    {
                        case DialogResult.Yes:
                            // Save record first  
                            SavePCsAndMacsRecord(PCsAndMacsRecord);
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

            // Create a new record
            PCsAndMacsRecord = new PCsAndMACsEditClass();

            // enable user edit window for populating the new record
            ClearPCsAndMACsEditPanel();

            // On New, render the PC / MAC panels
            PnlPCUser.Visible = true;
            PnlPCHardware.Visible = false; // This gets enabled when user selects the appropriate hardware template

        }

        private void RadPCActive_CheckedChanged(object sender, EventArgs e)
        {
            CmdSavePC.Enabled = true;
        }

        private void RadPCInactive_CheckedChanged(object sender, EventArgs e)
        {
            CmdSavePC.Enabled = true;
        }

        #endregion

        #endregion

        #region Servers

        #region Servers Declarations
        bool ServersInitialized = false;
        ListSortDirection ServersGridSortOrder = ListSortDirection.Ascending;
        int ServersGridSortColumn = 0;
        #endregion

        #region Servers Initialization
        private void PopulateServersGrid()
        {
            throw new NotImplementedException();
        }

        #endregion

        #endregion

        #region Documentation

        #region Documentation Declarations

        bool DocumentationInitialized = false;

        #endregion

        #region Documentation Initialization

        private void PopulateDocumentationGrid()
        {
            throw new NotImplementedException();
        }

        #endregion

        #endregion

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
#if false
            // Get the row index for whatever row was just clicked.  
            int row = e.RowIndex;
            if (row < 0)
            {
                // header was clicked (for sorting).  Toggle the sort order direction
                PCsAndMacsGridSortColumn = e.ColumnIndex;
                PCsAndMacsGridSortOrder = (GrdPCsAndMacs.Columns[PCsAndMacsGridSortColumn].HeaderCell.SortGlyphDirection == System.Windows.Forms.SortOrder.Descending) ? ListSortDirection.Ascending : ListSortDirection.Descending;
                return;
            }

            // If a new template entry has been created but not saved, prompt the user before populating the 
            //   edit area from the grid.

            if (PCsAndMacsRecord != null)
            {
                if (!PCsAndMacsRecord.IsSaved)
                {
                    DialogResult dlgResult = MessageBox.Show("The current record in the edit window has NOT been saved; do you want to save it first?", "SAVE PC / MAC RECORD?", MessageBoxButtons.YesNoCancel);
                    switch (dlgResult)
                    {
                        case DialogResult.Yes:
                            // Save record first  
                            SavePCsAndMacsRecord(PCsAndMacsRecord);
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

            // Load the selected grid data into the edit template. 
            PCsAndMacsRecord = new PCsAndMACsEditClass();
            PnlPCUser.Visible = true;
            PnlPCHardware.Visible = true;
            DisplayPCsAndMACsRecord(PCsAndMacsRecord, (int)GrdPCsAndMacs.Rows[row].Cells["pcmac_id"].Value);
            PCsAndMacsRecord.IsNew = false;
#endif
        }
    }
}


