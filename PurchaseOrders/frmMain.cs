using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using BSGlobals;

namespace PurchaseOrders
{
    public partial class FrmMain : Form
    {
        #region Declarations
        const string JobName = "Purchase Orders";
        int LookbackInYears = 5;

        OrderItemClass SelectedOrderItem;
        VendorListClass CurrentVendors = new VendorListClass();
        OrderClass CurrentOrder;
        ActiveDirectory UserInfo = new ActiveDirectory();
        Spreadsheet POSpreadsheet;
        Spreadsheet ERSpreadsheet;
        VersionStatusBar StatusBar;

        #endregion

        #region Initialization
        public FrmMain()
        {
            InitializeComponent();

            // Get the current (logged-in) username.  It will be in the form DOMAIN\username

            MainMenuStrip.Renderer = new CustomMenuStripRenderer();

            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job starting", JobName);

            cmdCopyItem.Enabled = false;
            cmdCopyOrder.Enabled = false;
            CmdApplyToNewOrder.Visible = false;
            LblCopied.Visible = false;

            CmdPaste.Visible = false;
            CmdPaste.BackColor = Color.Transparent;
            CmdNext.Enabled = false;
            CmdPrev.Enabled = false;
            CmdSaveVendor.Enabled = false;
            CmdSavePO.Enabled = false;
            CmdExpenseReport.Enabled = false;

            MnuSaveOrder.Enabled = false;
            MnuCopyItem.Enabled = false;
            MnuPaste.Enabled = false;
            MnuCopyOrder.Enabled = false;

            PnlVendor.Visible = false;
            PnlOrderButtons.Visible = false;
            PnlOrderDetail.Visible = false;
            LblCreatingSpreadsheet.Visible = false;

            chkFilterByVendor.Checked = false;
            cmbFilterByVendor.Enabled = false;

            // Get configuration values.  Configuration objects will generate events, so after this section, disable the configuration SAVE button
            //   as thsee change events will have enabled it.

            bool lookbackokay = int.TryParse(Config.GetConfigurationKeyValue("Purchasing", "LookbackInYears"), out LookbackInYears);
            UdLookbackInYears.Value = LookbackInYears;

            CmdSaveConfiguration.Enabled = false;

            PopulateExistingOrderGrid("");
            PopulateVendorComboBox(CmbVendorName, "");
            PopulateVendorComboBox(cmbFilterByVendor, "");

            // Create vendor text change events
            TxtAddressLine1.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtAddressLine2.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtCity.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtState.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtZipCode.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtContact.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtTelephone.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtFax.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);
            TxtNewsAccount.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);

            // Create order text change events
            TxtDate.TextChanged += new System.EventHandler(this.OrderTxtBox_TextChanged);

            // Add status bar (2 segment default, with version)
            StatusBar = new VersionStatusBar(this);

        }
        #endregion

        #region Populate Combo Boxes
        private void PopulateChargeToComboBox(int rowNum)
        {
            using (SqlDataReader rdr = SQLQuery("Proc_Select_Departments"))
            {
                if (rdr.HasRows)
                {
                    DataGridViewComboBoxCell dgvCell = (DataGridViewComboBoxCell)GrdOrderDetails.Rows[rowNum].Cells["ChargeTo"];
                    DataTable dt = new DataTable();
                    dt.Load(rdr);
                    dgvCell.DataSource = dt;
                    dgvCell.DisplayMember = "Dept";
                }
            }
        }

        private void PopulateClassificationComboBox(int rowNum)
        {
            using (SqlDataReader rdr = SQLQuery("Proc_Select_Classifications"))
            {
                if (rdr.HasRows)
                {
                    DataGridViewComboBoxCell dgvCell = (DataGridViewComboBoxCell)GrdOrderDetails.Rows[rowNum].Cells["Classification"];
                    DataTable dt = new DataTable();
                    dt.Load(rdr);
                    dgvCell.DataSource = dt;
                    dgvCell.DisplayMember = "Class";
                }
            }
        }

        private void PopulateExistingOrderGrid(string vendorFilter)
        {
            // Populate the existing order grid from SQL

            //GrdExistingOrders.DataSource = ""; // TBD This wipes out our named columns
            SqlParameter[] VendorParams = new SqlParameter[2];
            VendorParams[0] = new SqlParameter("@pvchrVendorName", vendorFilter);
            VendorParams[1] = new SqlParameter("@pvintLookbackInYears", LookbackInYears);
            using (SqlDataReader rdr = SQLQuery("Proc_Select_All_Orders_By_Vendorname", VendorParams))
            {
                //if (rdr.HasRows)
                {
                    GrdExistingOrders.Visible = true;
                    DataTable dt = new DataTable();
                    dt.Load(rdr);
                    GrdExistingOrders.DataSource = dt;
                }
            }
        }

        private void PopulateVendorComboBox(ComboBox combo, string vendorFilter)
        {
            // Populate the existing order grid from SQL

            SqlParameter[] VendorParams = new SqlParameter[1];
            VendorParams[0] = new SqlParameter("@pvchrVendorName", vendorFilter);
            using (SqlDataReader rdrVendor = SQLQuery("Proc_Select_Vendor", VendorParams))
            //            using (SqlDataReader rdr = DataIO.ExecuteQuery(
            //                Enums.DatabaseConnectionStringNames.Purchasing,
            //                CommandType.Text,
            //                "SELECT Vendor FROM tblVendors WHERE (Vendor <> '' AND Vendor is not NULL) ORDER BY Vendor"))
            {
                if (rdrVendor.HasRows)
                {
                    DataTable dt = new DataTable();
                    dt.Load(rdrVendor);
                    combo.DataSource = dt;
                    combo.DisplayMember = "Vendor";
                }
                combo.SelectedIndex = -1;
            }
        }
        #endregion

        #region PO and Vendor Record Functions

        /// <summary>
        /// Creates a new Purchase Order record in the tblOrders table
        /// </summary>
        /// <returns></returns>
        private OrderClass CreateNewPORecord()
        {
            // Because this is a shared database we need to create a record - that is, locking the PO# - as soon as
            //   the user starts the PO process.  This number is auto-incremented so it will never be reused,
            //   even if the user decides not to save the PO.

            OrderClass order = new OrderClass();
            try
            {
                using (SqlDataReader rdrPONum = SQLQuery("Proc_Insert_Order"))
                {
                    rdrPONum.Read();
                    {
                        // Update the PO# on the main page.
                        SafeTextBox(TxtPONumber, rdrPONum, "OrdID");
                        bool PONumOkay = int.TryParse(TxtPONumber.Text, out int PONum);
                        order.UpdatePONumber(PONum);
                        LblPONumber.Text = "Purchase Order D" + PONum.ToString("D5");

                        // Disable the SAVE button (and menu) items until the user actually enters some data.
                        CmdSavePO.Enabled = false; // This is a new (empty PO) so disable the SAVE button
                        MnuSaveOrder.Enabled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.ERROR, "Unable to create new purchase order:  " + ex.ToString(), JobName);
                MessageBox.Show("ERROR - " + "Unable to create new purchase order:  " + ex.ToString());
                return (null);
            }
            return (order);
        }

        private void AutoCreateNewPORecord(OrderItemClass selectedOrderItem)
        {
            // Create a new PO from the row that was selected from the existing PO List
            // Only a single row (row 0) is created if selectedOrderItem.SingleItemOnly is set
            // Note that we are forced to create an order record here to prevent anyone else from using this PO Number.
            //   If we don't complete the PO order we can delete it at the end.

            CurrentOrder = CreateNewPORecord();

            // Get all vendor info associated with this item

            SqlParameter[] OrderParams = new SqlParameter[2];
            OrderParams[0] = new SqlParameter("@pvintOrderID", SelectedOrderItem.SelectedOrderID);
            OrderParams[1] = new SqlParameter("@pvintLookbackInYears", LookbackInYears);
            using (SqlDataReader rdrOrder = SQLQuery("Proc_Select_All_Orders", OrderParams))
            {
                rdrOrder.Read(); // Read just the first record
                {
                    // Put order into order object and display

                    CurrentOrder.LoadOrderFromSQL(rdrOrder);
                    DisplayOrder(CurrentOrder);

                    // Put vendor into vendor object and display

                    CurrentVendors = new VendorListClass();
                    VendorClass v = new VendorClass(rdrOrder, UserInfo.Username);
                    CurrentVendors.VendorList.Add(v);
                    CurrentVendors.VendorInfoChanged = false;

                    string vendorname = CurrentVendors.GetCurrentVendorName();
                    SelectVendor(vendorname);

                }
            }

            // Put order details into order object
            int ItemRecordID = (SelectedOrderItem.SingleItemOnly) ? SelectedOrderItem.SelectItemRecordID : 0;
            SqlParameter[] ItemParams = new SqlParameter[2];
            ItemParams[0] = new SqlParameter("@pvintOrderID", SelectedOrderItem.SelectedOrderID);
            ItemParams[1] = new SqlParameter("@pvintItemRecordID", ItemRecordID);
            using (SqlDataReader rdrItem = SQLQuery("Proc_Select_Order_Item", ItemParams))
            {
                int row = 0;
                GrdOrderDetails.Rows.Clear();

                while (rdrItem.Read())
                {
                    GrdOrderDetails.Rows.Add();
                    PopulateChargeToComboBox(row);
                    PopulateClassificationComboBox(row);
                    DisplayLineItems(GrdOrderDetails.Rows[row], rdrItem);
                    CurrentOrder.LoadLineItemsFromSQL(rdrItem, true);
                    if (selectedOrderItem.SingleItemOnly) break; // If only a single entry was copied, exit the loop after a single iteration
                    row++;
                }
            }

            TxtPOTotal.Text = CurrentOrder.ComputeOrderTotal().ToString("C2");
        }

        private void SavePORecord(OrderClass currentOrder)
        {
            // Save all Order information in the current order record (pointed to by OrderID)
            OrderClass o = currentOrder;
            VendorClass v = CurrentVendors.VendorList[CurrentVendors.SelectedListIndex];

            if (o.PONumber > 0)
            {
                SqlParameter[] OrderParams = new SqlParameter[19];
                OrderParams[0] = new SqlParameter("@pvintOrdID", o.PONumber);
                OrderParams[1] = new SqlParameter("@pvchrVendor", v.VendorName);
                OrderParams[2] = new SqlParameter("@pvchrAdd1", v.AddrLine1);
                OrderParams[3] = new SqlParameter("@pvchrAdd2", v.AddrLine2);
                OrderParams[4] = new SqlParameter("@pvchrCity", v.City);
                OrderParams[5] = new SqlParameter("@pvchrState", v.State);
                OrderParams[6] = new SqlParameter("@pvchrZip", v.Zip);
                OrderParams[7] = new SqlParameter("@pvchrContact", v.Contact);
                OrderParams[8] = new SqlParameter("@pvchrPhone", v.Phone);
                OrderParams[9] = new SqlParameter("@pvchrFax", v.Fax);
                OrderParams[10] = new SqlParameter("@pvchrAcctNum", v.AcctNum);
                OrderParams[11] = new SqlParameter("@pvchrRefNum", o.OrderReference);
                OrderParams[12] = new SqlParameter("@pvchrDept", o.Department);
                OrderParams[13] = new SqlParameter("@pvchrExt", o.DeliverToPhone);
                OrderParams[14] = new SqlParameter("@pvdatOrdDate", o.OrderDate);
                OrderParams[15] = new SqlParameter("@pvchrDelTo", o.DeliverTo);
                OrderParams[16] = new SqlParameter("@pvchrTerms", o.Terms);
                OrderParams[17] = new SqlParameter("@pvchrComments", o.Comments);
                OrderParams[18] = new SqlParameter("@pvchrOwner", UserInfo.Username);
                SQLProcCall("Proc_Update_Order", OrderParams);

                SaveLineItems(currentOrder);
                currentOrder.OrderIsSaved = true;
                currentOrder.SaveCount++;
            }
            else
            {
                MessageBox.Show("Error: PO Number was zero on Save", "PO # NOT SET", MessageBoxButtons.OK);
            }
        }

        private void SaveLineItems(OrderClass currentOrder)
        {
            //  Delete any previous order items associated with this order ID

            SqlParameter[] DeleteParams = new SqlParameter[1];
            DeleteParams[0] = new SqlParameter("@pvintOrdID", CurrentOrder.PONumber);
            SQLProcCall("Proc_Delete_Order_Items", DeleteParams);

            // For each order item on the order item list, create (or update) the corresponding SQL record
            for (int i = 0; i < currentOrder.NumLineItems; i++)
            {
                LineItemsClass o = currentOrder.GetLineItems(i);
                SqlParameter[] OrderParams = new SqlParameter[11];
                OrderParams[0] = new SqlParameter("@pvintOrdID", currentOrder.PONumber);
                OrderParams[1] = new SqlParameter("@pvintRecID", o.OrderItemID);
                OrderParams[2] = new SqlParameter("@pvintQty", o.Quantity);
                OrderParams[3] = new SqlParameter("@pvchrUnits", o.Units);
                OrderParams[4] = new SqlParameter("@pvchrDescription", o.Description);
                OrderParams[5] = new SqlParameter("@pvcurUnitPrice", o.UnitPrice);
                OrderParams[6] = new SqlParameter("@pvchrChargeTo", o.ChargeTo);
                OrderParams[7] = new SqlParameter("@pvchrPurpose", o.Purpose);
                OrderParams[8] = new SqlParameter("@pvchrClass", o.Classification);
                OrderParams[9] = new SqlParameter("@pvchrTaxable", (o.IsTaxable ? "1" : "0"));
                OrderParams[10] = new SqlParameter("@pvchrOwner", UserInfo.Username);
                SqlDataReader rdrItem = SQLQuery("Proc_Update_Order_Item", OrderParams);
                try
                {
                    rdrItem.Read();
                    o.OrderItemID = (int)GetSQLValue(rdrItem, "RecID");
                }
                catch (Exception ex)
                {
                    DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.ERROR, "Unable to obtain the record ID for the updated order:  " + ex.ToString(), JobName);
                    MessageBox.Show("ERROR - " + "Unable to obtain the record ID for the updated order:  " + ex.ToString());
                }
            }
        }

        private VendorClass CreateNewVendorRecord()
        {
            try
            {
                using (SqlDataReader rdrVendor = SQLQuery("Proc_Insert_Vendor"))
                {
                    rdrVendor.Read();
                    {
                        VendorClass vendor = new VendorClass(rdrVendor, UserInfo.Username);
                        return (vendor);
                    }
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.ERROR, "Unable to create new vendor record:  " + ex.ToString(), JobName);
                MessageBox.Show("ERROR - " + "Unable to create new vendor record:  " + ex.ToString());
                return (null);
            }
        }

        private int SelectVendor(string vendorname)
        {
            int VendorID = 0;
            SqlParameter[] VendorParams = new SqlParameter[1];
            VendorParams[0] = new SqlParameter("@pvchrVendorName", vendorname);
            using (SqlDataReader rdrVendor = SQLQuery("Proc_Select_Vendor", VendorParams))
            {
                if (rdrVendor.HasRows)
                {
                    CurrentVendors.VendorList.Clear();
                    while (rdrVendor.Read())
                    {
                        VendorClass v = new VendorClass(rdrVendor, UserInfo.Username);
                        CurrentVendors.VendorList.Add(v);
                    }

                    // display the first item of the list here.
                    CurrentVendors.SelectedListIndex = 0;
                    DisplayVendor(CurrentVendors);
                    VendorID = CurrentVendors.VendorList[CurrentVendors.SelectedListIndex].VendorID;

                    // If more than one item in the list, enable the NEXT button
                    CmdPrev.Enabled = false;
                    CmdPrev.BackColor = Color.Transparent;
                    if (CurrentVendors.VendorList.Count > 1)
                    {
                        CmdNext.Enabled = true;
                        CmdNext.BackColor = Color.Yellow;
                    }
                    else
                    {
                        CmdNext.Enabled = false;
                        CmdNext.BackColor = Color.Transparent;
                    }
                }
            }
            return (VendorID);
        }

        private void SaveVendorRecord(VendorListClass currentVendors)
        {
            VendorClass v = currentVendors.VendorList[currentVendors.SelectedListIndex];

            if (v.VendorID > 0)
            {
                SqlParameter[] VendorParams = new SqlParameter[12];
                VendorParams[0] = new SqlParameter("@pvintVenID", v.VendorID);
                VendorParams[1] = new SqlParameter("@pvchrVendor", v.VendorName);
                VendorParams[2] = new SqlParameter("@pvchrAdd1", v.AddrLine1);
                VendorParams[3] = new SqlParameter("@pvchrAdd2", v.AddrLine2);
                VendorParams[4] = new SqlParameter("@pvchrCity", v.City);
                VendorParams[5] = new SqlParameter("@pvchrState", v.State);
                VendorParams[6] = new SqlParameter("@pvchrZip", v.Zip);
                VendorParams[7] = new SqlParameter("@pvchrContact", v.Contact);
                VendorParams[8] = new SqlParameter("@pvchrPhone", v.Phone);
                VendorParams[9] = new SqlParameter("@pvchrFax", v.Fax);
                VendorParams[10] = new SqlParameter("@pvchrAcctNum", v.AcctNum);
                VendorParams[11] = new SqlParameter("@pvchrOwner", v.Username);
                SQLProcCall("Proc_Update_Vendor", VendorParams);
            }
            else
            {
                MessageBox.Show("Error: Vendor ID was zero on Save", "VendorID NOT SET", MessageBoxButtons.OK);
            }
        }
        #endregion

        #region Data Display Functions
        private void DisplayOrder(OrderClass currentOrder)
        {
            // Displays the current order
            try
            {
                OrderClass o = currentOrder;
                SafeTextBox(TxtDate, o.OrderDate.Value.ToShortDateString());
                SafeTextBox(TxtDeliverto, o.DeliverTo);
                SafeTextBox(TxtDeliverToPhone, o.DeliverToPhone);
                SafeTextBox(TxtDepartment, o.Department);
                SafeTextBox(TxtTerms, o.Terms);
                SafeTextBox(TxtOrderReference, o.OrderReference);
                SafeTextBox(TxtComments, o.Comments);

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR - Unable to display current order:  " + ex.ToString());
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to display current order:  " + ex.ToString(), JobName);
            }
        }

        private void DisplayLineItems(DataGridViewRow dgvr, SqlDataReader rdrItem)
        {
            // Displays the current order's line items on the grid
            try
            {
                SafeTextBox(dgvr.Cells["Qty"], rdrItem, "Qty");
                SafeTextBox(dgvr.Cells["Units"], rdrItem, "Units");
                SafeTextBox(dgvr.Cells["PartNumber"], rdrItem, "Description");
                SafeTextBox(dgvr.Cells["ItemUnitPrice"], rdrItem, "UnitPrice");
                SafeTextBox(dgvr.Cells["Purpose"], rdrItem, "Purpose");
                SafeTextBox((DataGridViewComboBoxCell)dgvr.Cells["ChargeTo"], rdrItem, "ChargeTo");
                SafeTextBox((DataGridViewComboBoxCell)dgvr.Cells["Classification"], rdrItem, "Class");
                SafeTextBox((DataGridViewCheckBoxCell)dgvr.Cells["Taxable"], rdrItem, "Taxable");

                string s = "";
                try
                {
                    s = dgvr.Cells["Qty"].Value.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR - Unable to populate grid Quantity value from SQL:  " + ex.ToString());
                    DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate grid Quantity value from SQL:  " + ex.ToString(), JobName);
                }

                string t = "";
                try
                {
                    t = dgvr.Cells["ItemUnitPrice"].Value.ToString();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR - Unable to populate grid Unit Price value from SQL:  " + ex.ToString());
                    DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate grid Unit Price value from SQL:  " + ex.ToString(), JobName);
                }
                bool qtyokay = double.TryParse(s, out double qty);
                bool priceokay = double.TryParse(t, out double price);
                if (qtyokay && priceokay)
                {
                    dgvr.Cells["ItemTotalPrice"].Value = qty * price;
                }

                GrdOrderDetails.Columns["ItemUnitPrice"].DefaultCellStyle.Format = "c";
                GrdOrderDetails.Columns["ItemTotalPrice"].DefaultCellStyle.Format = "c";

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR trying display current line items:  " + ex.ToString());
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to correctly display current line items:  " + ex.ToString(), JobName);
            }
        }

        private void DisplayVendor(VendorListClass currentVendors)
        {
            // Displays everything but the VendorID and Owner

            try
            {
            VendorClass v = currentVendors.VendorList[currentVendors.SelectedListIndex];
            CurrentVendors.DisableSelectionEvent = true;
            SafeTextBox(CmbVendorName, v.VendorName);  // We DON'T want a selected_change event to happen here, so disable the event
            CurrentVendors.DisableSelectionEvent = false;

            SafeTextBox(TxtAddressLine1, v.AddrLine1);
            SafeTextBox(TxtAddressLine2, v.AddrLine2);
            SafeTextBox(TxtCity, v.City);
            SafeTextBox(TxtState, v.State);
            SafeTextBox(TxtZipCode, v.Zip);
            SafeTextBox(TxtContact, v.Contact);
            SafeTextBox(TxtTelephone, v.Phone);
            SafeTextBox(TxtFax, v.Fax);
            SafeTextBox(TxtNewsAccount, v.AcctNum);

            // As this was a new insert onto the Vendor panel, 
            //   reset the "VendorInfoChanged" flag as no data mods to the Vendor record have actually occurred.
            CurrentVendors.VendorInfoChanged = false;
            CurrentVendors.VendorIsSupplied = true;
            RenderVendorSaveButton(CurrentVendors);
            RenderSaveOrderButton();

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR trying display current vendor:  " + ex.ToString());
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to correctly display current vendor:  " + ex.ToString(), JobName);
            }
        }

        private void ClearPanelData()
        {
            // This is invoked when a new order is generated (clears any previous orders)

            OrderClass o = new OrderClass();
            DisplayOrder(o);

            VendorListClass v = new VendorListClass();
            v.VendorList.Add(new VendorClass());
            DisplayVendor(v);
            CurrentVendors.VendorIsSupplied = false;
            GrdOrderDetails.Rows.Clear();
        }
        #endregion

        #region Button Rendering Functions
        private void RenderVendorSaveButton(VendorListClass currentVendors)
        {
            if (currentVendors.VendorInfoChanged)
            {
                CmdSaveVendor.Enabled = true;
                CmdSaveVendor.BackColor = Color.Yellow;
            }
            else
            {
                CmdSaveVendor.Enabled = false;
                CmdSaveVendor.BackColor = Color.Transparent;
            }

        }

        private void RenderSaveOrderButton()
        {
            // SAVE PO button is not enabled unless we have both a vendor AND at least one line of detail
            CmdSavePO.Enabled = ((GrdOrderDetails.Rows.Count > 1) && (CurrentVendors.VendorIsSupplied));
            MnuSaveOrder.Enabled = CmdSavePO.Enabled;
        }

        private void RenderNextPrevButtons(VendorListClass CurrentVendors)
        {
            if (CurrentVendors.VendorList.Count > 1)
            {
                if (CurrentVendors.SelectedListIndex < CurrentVendors.VendorList.Count - 1)
                {
                    CmdNext.Enabled = true;
                    CmdNext.BackColor = Color.Yellow;
                }
                else
                {
                    CmdNext.Enabled = false;
                    CmdNext.BackColor = Color.Transparent;
                }

                if (CurrentVendors.SelectedListIndex > 0)
                {
                    CmdPrev.Enabled = true;
                    CmdPrev.BackColor = Color.Yellow;
                }
                else
                {
                    CmdPrev.Enabled = false;
                    CmdPrev.BackColor = Color.Transparent;
                }
            }
            else
            {
                CmdNext.Enabled = false;
                CmdNext.BackColor = Color.Transparent;
                CmdPrev.Enabled = false;
                CmdPrev.BackColor = Color.Transparent;
            }
        }
        #endregion

        #region Safe Value Assignment Functions

        // Functions to assign a value directly from a SQL field (or a string) into a control's text,
        //   with error capture.

        private void SafeTextBox(DataGridViewCheckBoxCell t, SqlDataReader rdr, string s)
        {
            try
            {
                t.Value = (rdr[s].ToString() == "0" ? 0 : -1);
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate grid checkbox cell from SQL :  " + ex.ToString(), JobName);
                t.Value = 0;
            }
        }

        private void SafeTextBox(DataGridViewComboBoxCell t, SqlDataReader rdr, string s)
        {
            try
            {
                t.Value = rdr[s].ToString();
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate grid combobox cell from SQL :  " + ex.ToString(), JobName);
                t.Value = "";
            }
        }

        private void SafeTextBox(DataGridViewCell t, SqlDataReader rdr, string s)
        {
            try
            {
                t.Value = rdr[s].ToString();
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate grid text cell from SQL :  " + ex.ToString(), JobName);
                t.Value = "";
            }
        }

        private void SafeTextBox(TextBox t, SqlDataReader rdr, string s)
        {
            try
            {
                t.Text = rdr[s].ToString();
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate textbox from SQL :  " + ex.ToString(), JobName);
                t.Text = "";
            }
        }

        private void SafeTextBox(ComboBox t, SqlDataReader rdr, string s)
        {
            try
            {
                t.Text = rdr[s].ToString();
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate combobox from SQL:  " + ex.ToString(), JobName);
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

        private double SaveTextBoxDouble(SqlDataReader rdr, string s)
        {
            try
            {
                string a = rdr[s].ToString();
                bool aokay = Double.TryParse(a, out double v);
                return (aokay ? v : 0);
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to populate double from SQL:  " + ex.ToString(), JobName);
                return 0;
            }
        }

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
        #endregion

        #region Timer-related Functions
        private void RenderStatusMsg(string s, bool enable)
        {
            if (enable)
            {
                timStatus.Interval = 3000;
                timStatus.Enabled = true;

                LblStatus.Text = s;
                LblStatus.Refresh();
                LblStatus.Visible = true;
            }
            else
            {
                LblStatus.Visible = false;
                timStatus.Enabled = false;
            }
        }
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

    }
}
