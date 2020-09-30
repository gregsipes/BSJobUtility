using System;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using BSGlobals;

namespace PurchaseOrders
{
    public partial class frmMain
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
            //TBD TBD TBD throw new NotImplementedException();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdSavePO_Click(object sender, EventArgs e)
        {
            RenderStatusMsg("Saving Order...", true);
            // Save the purchase order and vendor data
            SaveVendorRecord(CurrentVendors);
            CurrentOrder.UpdateOrderRecord(
                TxtDate.Text, TxtDeliverto.Text, TxtDeliverToPhone.Text, TxtDepartment.Text, TxtTerms.Text, TxtOrderReference.Text, TxtComments.Text);
            SavePORecord(CurrentOrder);
            RenderStatusMsg("Order Saved", true);
            CmdExpenseReport.Enabled = true;
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

            // Don't do anything if we don't have an order

            if (CurrentOrder != null)
            {
                LblCreatingSpreadsheet.Visible = true;
                LblStatus.Visible = false;
                POSpreadsheet = CreatePurchaseOrderSpreadsheet(CurrentOrder);
                LblCreatingSpreadsheet.Visible = false;
                POSpreadsheet.Show();
            }
            else
            {
                MessageBox.Show("Try creating a PO before trying to print/preview");
            }
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
            // Don't do anything if we don't have an order

            if (CurrentOrder != null)
            {
                LblCreatingSpreadsheet.Visible = true;
                LblStatus.Visible = false;
                POSpreadsheet = CreatePurchaseOrderSpreadsheet(CurrentOrder);
                LblCreatingSpreadsheet.Visible = false;
                POSpreadsheet.Print();
            }
            else
            {
                MessageBox.Show("Try creating a PO before trying to print/preview");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdExpenseReport_Click(object sender, EventArgs e)
        {
            // Generate an expense report from the standard expense report template
            LblCreatingSpreadsheet.Visible = true;
            LblStatus.Visible = false;
            ERSpreadsheet = CreateExpenseReportSpreadsheet(CurrentOrder);
            LblCreatingSpreadsheet.Visible = false;
            ERSpreadsheet.Show();
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmbFilterByVendor_SelectedIndexChanged(object sender, EventArgs e)
        {
            string vendorname = cmbFilterByVendor.Text;
            // Redisplay the vendor grid with records only from the specified vendor
            PopulateExistingOrderGrid(vendorname);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdNext_Click(object sender, EventArgs e)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdPrev_Click(object sender, EventArgs e)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VendorTxtBox_TextChanged(object sender, EventArgs e)
        {
            CurrentVendors.VendorInfoChanged = true;
            RenderVendorSaveButton(CurrentVendors);
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

            SaveVendorRecord(CurrentVendors);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdNewVendor_Click(object sender, EventArgs e)
        {
            // Create a new vendor
            VendorListClass v = new VendorListClass();
            v.VendorList.Add(CreateNewVendorRecord());
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdCopyOrder_Click(object sender, EventArgs e)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdNewOrder_Click(object sender, EventArgs e)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdPaste_Click(object sender, EventArgs e)
        {
            // If an existing record was already selected (better have been!) then paste it into a new PO

            AutoCreateNewPORecord(SelectedOrderItem);
            PnlOrderDetail.Visible = true;
            PnlOrderButtons.Visible = true;
            PnlVendor.Visible = true;
            CmdPaste.BackColor = Color.Transparent;
            LblCopied.Visible = false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdApplyToNewOrder_Click(object sender, EventArgs e)
        {
            // Apply a copied ordr to a new order
            CmdApplyToNewOrder.Visible = false;
            CmdPaste_Click(sender, e);
            CmdPaste.Visible = false;
            TabOrders.SelectTab("TabNewOrders");
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
            // Enable Copy buttons
            cmdCopyItem.Enabled = true;
            cmdCopyOrder.Enabled = true;
            // e.ColumnIndex and e.RowIndex contain the column and row values, respectively (these are zero-based)
            SelectedOrderItem = new OrderItemClass(e.RowIndex, e.ColumnIndex);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdOrderDetails_CellEndEdit(object sender, DataGridViewCellEventArgs e)
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
                    catch (Exception ex)
                    {

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
                    catch (Exception ex)
                    {
                        //tbd
                    }
                    priceokay = double.TryParse(t, out double price);
                    d.UnitPrice = (priceokay ? price : 0);
                    break;
                case 4: // Total Price - Computed value, never entered
                    break;
                case 5: // Taxable
                    d.IsTaxable = (bool)GrdOrderDetails.Rows[row].Cells[5].Value;
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
            RenderSaveOrderButton();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GrdOrderDetails_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
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
            TabOrders.SelectTab("TabConfiguration");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UdLookbackInYears_ValueChanged(object sender, EventArgs e)
        {
            CmdSaveConfiguration.Enabled = true;
            LookbackInYears = (int)UdLookbackInYears.Value;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CmdSaveConfiguration_Click(object sender, EventArgs e)
        {
            CmdSaveConfiguration.Enabled = false;

            // Here is where we save any and all configuration values

            Config.SetConfigurationKeyValue("Purchasing", "LookbackInYears", UdLookbackInYears.Value.ToString());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChkFilterByVendor_CheckedChanged(object sender, EventArgs e)
        {
            cmbFilterByVendor.Enabled = chkFilterByVendor.Checked;
            if (!chkFilterByVendor.Checked)
            {
                PopulateVendorComboBox(cmbFilterByVendor, "");
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

        #endregion

        #region Timer Events

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimStatus_Tick(object sender, EventArgs e)
        {
            RenderStatusMsg("", false);
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
            Application.Exit();
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
                                // TBD
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

        #endregion

    }
}
