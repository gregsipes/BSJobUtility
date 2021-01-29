using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using BSGlobals;

namespace PurchaseOrders
{
    public partial class FrmMain
    {
        // ALL Events are put here so that we don't clutter the main code with too much stuff.

        #region Menu Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuCopyItem_Click(object sender, EventArgs e)
        {
            CmdCopyItem_Click(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuCopyOrder_Click(object sender, EventArgs e)
        {
            CmdCopyOrder_Click(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuPaste_Click(object sender, EventArgs e)
        {
            CmdPaste_Click(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuNewOrder_Click(object sender, EventArgs e)
        {
            CmdNewOrder_Click(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuExit_Click(object sender, EventArgs e)
        {
            CmdExit_Click(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuSaveOrder_Click(object sender, EventArgs e)
        {
            CmdSavePO_Click(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuPrintPreview_Click(object sender, EventArgs e)
        {
            CmdPrintPreview_Click(sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MnuPrint_Click(object sender, EventArgs e)
        {
            CmdPrint_Click(sender, e);
        }

        #endregion

        #region Order Control Events
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OrderTxtBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //TBD TBD TBD throw new NotImplementedException();
            }
            catch (Exception ex)
            {
                BroadcastError("Error in OrderTxtBox text changed event: ", ex);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdSavePO_Click(object sender, EventArgs e)
        {
            try
            {
                RenderStatusMsg("Saving Order...", true);
                // Save the purchase order and vendor data
                SaveVendorRecord(CurrentVendors);
                CurrentOrder.UpdateOrderRecord(
                    TxtDate.Text, TxtDeliverto.Text, TxtDeliverToPhone.Text, TxtDepartment.Text, TxtTerms.Text, TxtOrderReference.Text, TxtComments.Text, UserInfo.ADUserList[0].FirstName + " " + UserInfo.ADUserList[0].LastName);
                SavePORecord(CurrentOrder);
                RenderStatusMsg("Order Saved", true);
                CmdExpenseReport.Enabled = true;
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to save PO: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdPrintPreview_Click(object sender, EventArgs e)
        {
            // Create a spreadsheet and display it
            // NOTE: The spreadsheet code is in file Spreadsheets.cs

            PrintSpreadsheet(CurrentOrder, CurrentVendors.VendorList[CurrentVendors.SelectedListIndex], true, sender, e);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdPrint_Click(object sender, EventArgs e)
        {
            // Print to the default printer
            // NOTE: The spreadsheet code is in file Spreadsheets.cs
            //  TBD the behavior of PrintPreview is such that if you close the copy of the spreadsheet being previewed,
            //    it is lost and has to be re-created here when printing.

            PrintSpreadsheet(CurrentOrder, CurrentVendors.VendorList[CurrentVendors.SelectedListIndex], false, sender, e);
        }

        private void PrintSpreadsheet(OrderClass currentOrder, VendorClass vendor, bool preview, object sender, EventArgs e)
        {
            try
            {
                // Don't do anything if we don't have an order
                if (currentOrder != null)
                {
                    // Order must be saved first

                    if (!currentOrder.OrderIsSaved)
                    {
                        DialogResult result = MessageBox.Show("Order must be saved before previewing:  Save now?", "ORDER NOT SAVED", MessageBoxButtons.YesNo);
                        if (result != DialogResult.Yes) return;
                        CmdSavePO_Click(sender, e); ;
                    }

                    LblCreatingSpreadsheet.Visible = true;
                    StatusBar.AddText(0, "Creating Spreadsheet");
                    LblStatus.Visible = false;
                    POSpreadsheet = CreatePurchaseOrderSpreadsheet(currentOrder, vendor);

                    LblCreatingSpreadsheet.Visible = false;
                    StatusBar.AddText(0, "");
                    if (preview)
                    {
                        POSpreadsheet.Show();
                    } 
                    else
                    {
                        POSpreadsheet.Print();
                    }
                }
                else
                {
                    MessageBox.Show("Try creating a PO before trying to print/preview");
                }
            }
            catch (Exception ex)
            {
                BroadcastError("Error during Print Preview event: ", ex);
            }
        }

        private void CmdPreviewArchive_Click(object sender, EventArgs e)
        {
            // Create an archive spreadsheet and display it
            // NOTE: The spreadsheet code is in file Spreadsheets.cs

            try
            {
                DataGridViewRow ArchiveRow = GrdExistingOrders.Rows[SelectedOrderItem.SelectedRow];
                PrintArchiveSpreadsheet(ArchiveRow, true);
            }
            catch (Exception ex)
            {
                BroadcastError("Error - Unable to preview the selected row", ex);
            }
        }

        private void CmdPrintArchive_Click(object sender, EventArgs e)
        {
            // Create an archive spreadsheet and display it
            // NOTE: The spreadsheet code is in file Spreadsheets.cs

            try
            {
                DataGridViewRow ArchiveRow = GrdExistingOrders.Rows[SelectedOrderItem.SelectedRow];
                PrintArchiveSpreadsheet(ArchiveRow, false);
            }
            catch (Exception ex)
            {
                BroadcastError("Error - Unable to preview the selected row", ex);
            }
        }

        private void PrintArchiveSpreadsheet(DataGridViewRow archiveRow,  bool preview)
        {
            // Don't do anything if we don't have an order
            try
            {
                if (SelectedOrderItem != null)
                {
                    // We know that the SelectedOrderItem OrderItemClass object has been populated, which gives us a row number.
                    //   It contains neither order information nor vendor information.  We still need to get
                    //      Order Record
                    //      Vendor Info
                    //      Order Line Items

                    // Get the PO Number off the item grid
                    int PONumber = (int)archiveRow.Cells[0].Value;
                    OrderClass ArchivedOrder = new OrderClass();
                    ArchivedOrder.UpdatePONumber(PONumber);

                    // Load the PO
                    SqlParameter[] OrderParams = new SqlParameter[2];
                    OrderParams[0] = new SqlParameter("@pvintOrderID", PONumber);
                    OrderParams[1] = new SqlParameter("@pvintLookbackInYears", LookbackInYears);
                    using (SqlDataReader rdrOrder = SQLQuery("Proc_Select_All_Orders", OrderParams))
                    {
                        rdrOrder.Read(); // Read just the first record
                        ArchivedOrder.LoadOrderFromSQL(rdrOrder);
                        ArchivedOrder.OrderDate = Convert.ToDateTime(rdrOrder["OrdDate"]);
                        // Get Vendor info from the grid.  NOTE:  VendorID is NOT maintained within the order!!!!!

                        VendorClass vendor = new VendorClass(rdrOrder, UserInfo.Username)
                        {
                            VendorName = archiveRow.Cells[3].Value.ToString(), // Vendor Name
                            AddrLine1 = archiveRow.Cells[9].Value.ToString(),  // Addr Line 1
                            AddrLine2 = archiveRow.Cells[10].Value.ToString(), // Addr Line 2
                            City = archiveRow.Cells[11].Value.ToString(),      // City
                            State = archiveRow.Cells[12].Value.ToString(),     // State
                            Zip = archiveRow.Cells[13].Value.ToString(),       // Zip
                            Contact = archiveRow.Cells[14].Value.ToString(),   // Contact
                            Phone = archiveRow.Cells[15].Value.ToString(),     // Phone
                            Fax = archiveRow.Cells[16].Value.ToString(),       // Fax
                            AcctNum = archiveRow.Cells[17].Value.ToString(),   // AcctNum
                            Username = archiveRow.Cells[1].Value.ToString()    // Owner
                        };

                        // Get all PO Line Items associated with this order

                        SqlParameter[] ItemParams = new SqlParameter[2];
                        ItemParams[0] = new SqlParameter("@pvintOrderID", PONumber);
                        ItemParams[1] = new SqlParameter("@pvintItemRecordID", 0); // Indicates ALL items associated with this order
                        using (SqlDataReader rdrItem = SQLQuery("Proc_Select_Order_Item", ItemParams))
                        {
                            while (rdrItem.Read())
                            {
                                ArchivedOrder.LoadLineItemsFromSQL(rdrItem, true);
                            }
                        }

                        POSpreadsheet = CreatePurchaseOrderSpreadsheet(ArchivedOrder, vendor);
                        if (preview)
                        {
                            POSpreadsheet.Show();
                        }
                        else
                        {
                            POSpreadsheet.Print();
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Try creating a PO before trying to print/preview");
                }
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to preview an archive: ", ex);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdExpenseReport_Click(object sender, EventArgs e)
        {
            try
            {
                // Generate an expense report from the standard expense report template
                LblCreatingSpreadsheet.Visible = true;
                StatusBar.AddText(0, "Creating Spreadsheet");
                LblStatus.Visible = false;
                ERSpreadsheet = CreateExpenseReportSpreadsheet(CurrentOrder);
                LblCreatingSpreadsheet.Visible = false;
                StatusBar.AddText(0, "");
                ERSpreadsheet.Show();
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to generate an expense report: ", ex);
            }
        }

        #endregion

        #region Vendor Control Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmbVendorName_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // Update the vendor data with the new vendor selection
                //  Note that the vendor table may have duplicate vendor names depending on 
                //  other vendor content.  We need to pull in all records with the same name
                //  (Using LIKE perhaps)
                //  and then save the contents in a list so that we can cycle through them
                //  if the user so desires using the next/prev buttons.

                if (CurrentVendors.DisableSelectionEvent) return;

                // Note that Next/Prev buttons should be visible only if there are multiple vendor records.

                string vendorname = CmbVendorName.Text;
                if (vendorname.Length <= 0) return;

                SelectVendor(vendorname);
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to select Vendor: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmbFilterByVendor_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string vendorname = cmbFilterByVendor.Text;
                // Redisplay the vendor grid with records only from the specified vendor
                PopulateExistingOrderGrid(vendorname);
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to filter by vendor: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdNext_Click(object sender, EventArgs e)
        {
            try
            {
                if (CurrentVendors.SelectedListIndex < CurrentVendors.VendorList.Count - 1)
                {
                    CurrentVendors.SelectedListIndex++;
                    CurrentVendors.VendorInfoChanged = false;
                    RenderNextPrevButtons(CurrentVendors);
                    RenderVendorSaveButton(CurrentVendors);
                    DisplayVendor(CurrentVendors);
                }
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to go to the next vendor record: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdPrev_Click(object sender, EventArgs e)
        {
            try
            {
                if (CurrentVendors.SelectedListIndex > 0)
                {
                    CurrentVendors.SelectedListIndex--;
                    CurrentVendors.VendorInfoChanged = false;
                    RenderVendorSaveButton(CurrentVendors);
                    RenderNextPrevButtons(CurrentVendors);
                    DisplayVendor(CurrentVendors);
                }
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to go to the previous vendor record: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VendorTxtBox_TextChanged(object sender, EventArgs e)
        {
            try
            {
                CurrentVendors.VendorInfoChanged = true;
                RenderVendorSaveButton(CurrentVendors);
            }
            catch (Exception ex)
            {
                BroadcastError("Error while updating a text box on the vendor panel: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdSaveVendor_Click(object sender, EventArgs e)
        {
            // On Vendor Save, we need to either
            //   1) Update an existing vendor (if this was copied from a previous order)
            //   2) Populate a new vendor (which is hopefully ALSO just an update because the New Vendor button should have created the new vendor record for us)

            try
            {
                SaveVendorRecord(CurrentVendors);
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to save vendor record: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdNewVendor_Click(object sender, EventArgs e)
        {
            try
            {
                // Create a new vendor
                VendorListClass v = new VendorListClass();
                v.VendorList.Add(CreateNewVendorRecord());
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to create a new vendor: ", ex);
            }
        }

        #endregion

        #region Copy/Paste Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdCopyItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Take the record currently pointed to and copy it

                int ItemRecordID = (int)GrdExistingOrders.Rows[SelectedOrderItem.SelectedRow].Cells["ItemRecordID"].Value;
                int OrderID = (int)GrdExistingOrders.Rows[SelectedOrderItem.SelectedRow].Cells["PO"].Value;
                SelectedOrderItem.SelectItemRecordID = ItemRecordID;
                SelectedOrderItem.SelectedOrderID = OrderID;
                SelectedOrderItem.SingleItemOnly = true; // Only this one detail should be copied
                                                         // Enable the PASTE button on the New PO tab

                LblCopied.Visible = true;
                CmdPaste.Visible = true;
                CmdPaste.BackColor = Color.Yellow;
                CmdApplyToNewOrder.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to copy a archived line item: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdCopyOrder_Click(object sender, EventArgs e)
        {
            try
            {
                // Take the order record currently pointed to and copy it

                int ItemRecordID = (int)GrdExistingOrders.Rows[SelectedOrderItem.SelectedRow].Cells["ItemRecordID"].Value;
                int OrderID = (int)GrdExistingOrders.Rows[SelectedOrderItem.SelectedRow].Cells["PO"].Value;
                SelectedOrderItem.SelectItemRecordID = ItemRecordID;
                SelectedOrderItem.SelectedOrderID = OrderID;
                SelectedOrderItem.SingleItemOnly = false; // ALL order details to be copied
                                                          // Enable the PASTE button on the New PO tab

                LblCopied.Visible = true;
                CmdPaste.Visible = true;
                CmdPaste.BackColor = Color.Yellow;
                CmdApplyToNewOrder.Visible = true;
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to copy an archived purchase order: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdNewOrder_Click(object sender, EventArgs e)
        {
            try
            {
                // Create a new PO
                CurrentOrder = CreateNewPORecord();

                // On new order, render all panels visible
                ClearPanelData();
                PnlVendor.Visible = true;
                PnlOrderButtons.Visible = true;
                PnlOrderDetail.Visible = true;
                CmdPaste.Visible = false;
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to create a new PO: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdPaste_Click(object sender, EventArgs e)
        {
            try
            {
                // If an existing record was already selected (better have been!) then paste it into a new PO

                AutoCreateNewPORecord(CurrentOrder, SelectedOrderItem);
                PnlOrderDetail.Visible = true;
                PnlOrderButtons.Visible = true;
                PnlVendor.Visible = true;
                CmdPaste.BackColor = Color.Transparent;
                LblCopied.Visible = false;
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to paste the selected archived PO: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdApplyToNewOrder_Click(object sender, EventArgs e)
        {
            try
            {
                // Apply a copied order to a new order
                CmdApplyToNewOrder.Visible = false;
                CmdPaste_Click(sender, e);
                CmdPaste.Visible = false;
                TabOrders.SelectTab("TabNewOrders");
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to apply this item to a new order: ", ex);
            }
        }

        #endregion

        #region Grid Events
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdExistingOrders_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // Enable Copy buttons
                cmdCopyItem.Enabled = true;
                cmdCopyOrder.Enabled = true;
                // e.ColumnIndex and e.RowIndex contain the column and row values, respectively (these are zero-based)
                SelectedOrderItem = new OrderItemClass(e.RowIndex, e.ColumnIndex);
            }
            catch (Exception ex)
            {
                BroadcastError("Error while selecting an existing order: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdOrderDetails_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int row = e.RowIndex;
                if (row >= CurrentOrder.NumLineItems)
                {
                    CurrentOrder.AddDetailRecord();
                    PopulateChargeToComboBox(row);
                    PopulateClassificationComboBox(row);
                }
                // Check if this is either the Quantity or Unit Price cell.  If so,
                //   and if both cells are now filled, check that they are valid real numbers
                //   and compute the total price.

                int col = e.ColumnIndex;
                bool qtyokay = false;
                bool priceokay = false;
                LineItemsClass d = CurrentOrder.GetLineItems(row);
                switch (col)
                {
                    case 0: // Qty
                        string s = "";
                        try
                        {
                            s = GrdOrderDetails.Rows[row].Cells["Qty"].Value.ToString();
                        }
                        catch
                        {
                            MessageBox.Show("Invalid Quantity Value - Please re-enter");
                        }
                        qtyokay = int.TryParse(s, out int qty);
                        d.Quantity = (qtyokay ? qty : 0);
                        break;
                    case 1: // Units
                        d.Units = GrdOrderDetails.Rows[row].Cells[1].Value.ToString();
                        break;
                    case 2: // Description
                        d.Description = GrdOrderDetails.Rows[row].Cells[2].Value.ToString();
                        break;
                    case 3: // Unit Price
                        string t = "";
                        try
                        {
                            t = GrdOrderDetails.Rows[row].Cells[3].Value.ToString();
                        }
                        catch
                        {
                            MessageBox.Show("Invalid Unit Price Value - Please re-enter");
                        }
                        priceokay = double.TryParse(t, out double price);
                        d.UnitPrice = (priceokay ? price : 0);
                        break;
                    case 4: // Total Price - Computed value, never entered
                        break;
                    case 5: // Taxable
                        d.IsTaxable = (GrdOrderDetails.Rows[row].Cells[5].Value == null ? false : (bool)GrdOrderDetails.Rows[row].Cells[5].Value);
                        break;
                    case 6: // Charge To (will fail if ChargeTo combobox isn't populated)
                        d.ChargeTo = GrdOrderDetails.Rows[row].Cells[6].Value.ToString();
                        break;
                    case 7: // Classification
                        d.Classification = GrdOrderDetails.Rows[row].Cells[7].Value.ToString();
                        break;
                    case 8: // Purpose
                        d.Purpose = GrdOrderDetails.Rows[row].Cells[8].Value.ToString();
                        break;
                }

                GrdOrderDetails.Rows[row].Cells[4].Value = (double)d.Quantity * d.UnitPrice;
                CurrentOrder.UpdateRow(row, d);
                TxtPOTotal.Text = CurrentOrder.ComputeOrderTotal().ToString("C2");
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to modify the grid cell: ", ex);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdOrderDetails_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Ignore data entry errors on the grid (mainly related to null data items)
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdOrderDetails_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            // Since the user can add rows the row count is always at least 1.  Enable the SAVE button if the count exceeds 1
            try
            {
                RenderSaveOrderButton();
            }
            catch (Exception ex)
            {
                BroadcastError("Error while trying to add a row to the order details: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdOrderDetails_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (CurrentOrder is null) return; // nothing to do if no previous order has been entered.

            try
            {
                // Remove this line item from the current order
                int row = e.RowIndex;
                CurrentOrder.RemoveRow(row);

                // If there are no rows left then disable the SAVE button
                // Since the user can add rows the row count is always at least 1.  Enable the SAVE button if the count exceeds 1
                RenderSaveOrderButton();

                // Recompute all totals
                TxtPOTotal.Text = CurrentOrder.ComputeOrderTotal().ToString("C2");

                // Regen the line item records
                SaveLineItems(CurrentOrder);
            }
            catch (Exception ex)
            {
                BroadcastWarning("WARNING - Error trying to clear all rows from the current purchase order: ", ex);
            }

        }

        #endregion

        #region Configuration Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Open the configuration tab
            try
            {
                TabOrders.SelectTab("TabConfiguration");
            }
            catch (Exception ex)
            {
                BroadcastError("Error while selecting a tab menu item: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UdLookbackInYears_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                CmdSaveConfiguration.Enabled = true;
                LookbackInYears = (int)UdLookbackInYears.Value;
            }
            catch (Exception ex)
            {
                BroadcastError("Error while trying to update the Lookback value: ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdSaveConfiguration_Click(object sender, EventArgs e)
        {
            try
            {
                CmdSaveConfiguration.Enabled = false;

                // Here is where we save any and all configuration values

                Config.SetConfigurationKeyValue("Purchasing", "LookbackInYears", UdLookbackInYears.Value.ToString());
            }
            catch (Exception ex)
            {
                BroadcastError("Error while trying to save the new configuration: ", ex);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChkFilterByVendor_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                cmbFilterByVendor.Enabled = chkFilterByVendor.Checked;
                if (!chkFilterByVendor.Checked)
                {
                    PopulateVendorComboBox(cmbFilterByVendor, "");
                }
            }
            catch (Exception ex)
            {
                BroadcastError("Error while trying to select a vendor filter option: ", ex);
            }

        }
        #endregion

        #region Tab Control Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TabOrders_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Enable the "edit" menu items ONLY if we are on the Existing Orders tab

            try
            {
                int tabindex = TabOrders.SelectedIndex;
                if (TabOrders.TabPages[tabindex].Text == "Existing Orders")
                {
                    MnuCopyItem.Enabled = true;
                    MnuPaste.Enabled = true;
                    MnuCopyOrder.Enabled = true;
                }
                else
                {
                    MnuCopyItem.Enabled = false;
                    MnuPaste.Enabled = false;
                    MnuCopyOrder.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to select the orders tab: ", ex);
            }

        }

        #endregion

        #region Timer Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimStatus_Tick(object sender, EventArgs e)
        {
            //try
            {
                RenderStatusMsg("", false);
            }
            //catch (Exception ex)
            {
                //BroadcastError("", ex); // don't gen an error here or we could make it happen on every tick.
            }
        }

        #endregion

        #region Termination Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdExit_Click(object sender, EventArgs e)
        {
            try
            {
                Application.Exit();
            }
            catch (Exception ex)
            {
                BroadcastError("Error trying to exit this app!!! ", ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            // TBD Check if the order has been saved, prompt if it hasn't been.
            //   Save (if required), then close

            try
            {
                if (CurrentOrder != null)
                {
                    if (!CurrentOrder.OrderIsSaved)
                    {
                        DialogResult result = MessageBox.Show("This PO has not been saved.  Save it now?", "SAVE PURCHASE ORDER", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            CmdSavePO_Click(sender, e);
                        }
                        else
                        {
                            // If this PO was never saved in the first place (SaveCount = 0) then delete it from the tblOrders table
                            if (CurrentOrder.SaveCount == 0)
                            {
                                try
                                {
                                    SqlParameter[] OrderParams = new SqlParameter[1];
                                    OrderParams[0] = new SqlParameter("@pvintOrdID", CurrentOrder.PONumber);
                                    SQLProcCall("Proc_Delete_Order", OrderParams);
                                }
                                catch (Exception ex)
                                {
                                    DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.WARNING, "Unable to delete the unsaved purchase order from the Orders table on app termination: " + ex.ToString(), JobName);
                                }
                            }

                            // TBD We should ALSO delete a NEW VENDOR or ORDER record if it doesn't get updated
                            // TBD We should ALSO prompt the user if an updated vendor or order record isn't saved before exiting

                        }
                    }

                    DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job completed", JobName);
                }

                if (POSpreadsheet != null)
                {
                    POSpreadsheet.Terminate();
                }
                if (ERSpreadsheet != null)
                {
                    ERSpreadsheet.Terminate();
                }
            }
            catch (Exception ex)
            {
                BroadcastError("Error while trying to close this app!!! ", ex);
            }
        }

        #endregion

    }
}
