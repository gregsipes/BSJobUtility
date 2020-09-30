using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace PurchaseOrders
{
    public partial class frmMain
    {
        // ALL internal classes are put here so that we don't clutter the main code with too much stuff.

        #region VendorClass
        /// <summary>
        /// 
        /// </summary>
        private class VendorClass
        {
            public int VendorID { get; private set; }
            public string VendorName { get; set; }
            public string AddrLine1 { get; set; }
            public string AddrLine2 { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string Zip { get; set; }
            public string Contact { get; set; }
            public string Phone { get; set; }
            public string Fax { get; set; }
            public string AcctNum { get; set; }
            public string Username { get; set; }

            public VendorClass()
            {
                VendorID = 0;
                VendorName = "";
                AddrLine1 = "";
                AddrLine2 = "";
                City = "";
                State = "";
                Zip = "";
                Contact = "";
                Phone = "";
                Fax = "";
                AcctNum = "";
                Username = "";
            }

            public VendorClass(SqlDataReader rdrVendor, string username)
            {
                // If invoked from an Order record, 
                //   These records contain no vendor ID
                //   Fields AddrLine1 and AddrLine2 are named Add1 and Add2, respectively
                try { VendorID = (int)GetSQLValue(rdrVendor, "VenID"); } catch { VendorID = 0; }
                VendorName = (string)GetSQLValue(rdrVendor, "Vendor");
                try { AddrLine1 = (string)GetSQLValue(rdrVendor, "AddrLine1"); } catch { AddrLine1 = (string)GetSQLValue(rdrVendor, "Add1"); }
                try { AddrLine2 = (string)GetSQLValue(rdrVendor, "AddrLine2"); } catch { AddrLine2 = (string)GetSQLValue(rdrVendor, "Add2"); }
                City = (string)GetSQLValue(rdrVendor, "City");
                State = (string)GetSQLValue(rdrVendor, "State");
                Zip = (string)GetSQLValue(rdrVendor, "Zip");
                Contact = (string)GetSQLValue(rdrVendor, "Contact");
                Phone = (string)GetSQLValue(rdrVendor, "Phone");
                Fax = (string)GetSQLValue(rdrVendor, "Fax");
                AcctNum = (string)GetSQLValue(rdrVendor, "AcctNum");
                Username = username;
            }

            internal void UpdateVendorID(int vendorID)
            {
                VendorID = vendorID;
            }
        }
        #endregion

        #region VendorListClass
        /// <summary>
        /// 
        /// </summary>
        private class VendorListClass
        {
            public List<VendorClass> VendorList { get; set; }
            public int SelectedListIndex { get; set; }
            public bool VendorInfoChanged { get; set; }
            public bool VendorIsSupplied { get; set; }
            public bool DisableSelectionEvent { get; set; }

            public VendorListClass()
            {
                VendorList = new List<VendorClass>();
                SelectedListIndex = 0;
                VendorInfoChanged = false;
                VendorIsSupplied = false;
                DisableSelectionEvent = false;
            }

            internal string GetCurrentVendorName()
            {
                return (VendorList[SelectedListIndex].VendorName);
            }
        }
        #endregion

        #region OrderClass
        /// <summary>
        /// 
        /// </summary>
        private class OrderClass
        {
            private List<LineItemsClass> LineItems { get; set; }
            public int? PONumber { get; private set; }
            public DateTime? OrderDate { get; private set; }
            public string DeliverTo { get; private set; }
            public string DeliverToPhone { get; private set; }
            public string Department { get; private set; }
            public string Terms { get; private set; }
            public string OrderReference { get; private set; }
            public string Comments { get; private set; }
            public double OrderTotal { get; private set; }
            public bool OrderIsSaved { get; set; }
            public int SaveCount { get; set; }
            public OrderClass()
            {
                LineItems = new List<LineItemsClass>();
                OrderDate = DateTime.Now;
                OrderTotal = 0;
                OrderIsSaved = false;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="orderDate"></param>
            /// <param name="deliverTo"></param>
            /// <param name="deliverToPhone"></param>
            /// <param name="department"></param>
            /// <param name="terms"></param>
            /// <param name="orderReference"></param>
            /// <param name="comments"></param>
            public void UpdateOrderRecord(string orderDate, string deliverTo, string deliverToPhone, string department, string terms, string orderReference, string comments)
            {
                bool dateokay = DateTime.TryParse(orderDate, out DateTime dt);
                OrderDate = (dateokay ? dt : DateTime.Now);
                DeliverTo = deliverTo;
                DeliverToPhone = deliverToPhone;
                Department = department;
                Terms = terms;
                OrderReference = orderReference;
                Comments = comments;
                OrderIsSaved = false;
                SaveCount = 0;
            }

            /// <summary>
            /// 
            /// </summary>
            internal double ComputeOrderTotal()
            {
                OrderTotal = 0;
                for (int i = 0; i < LineItems.Count; i++)
                {
                    OrderTotal += LineItems[i].TotalPrice;
                }
                return (OrderTotal);
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="currentOrder"></param>
            /// <param name="rdrOrder"></param>
            public void LoadOrderFromSQL(SqlDataReader rdrOrder)
            {
                OrderDate = DateTime.Now;
                DeliverTo = (string)GetSQLValue(rdrOrder, "DelTo");
                DeliverToPhone = (string)GetSQLValue(rdrOrder, "Ext");
                Department = (string)GetSQLValue(rdrOrder, "Dept");
                Terms = (string)GetSQLValue(rdrOrder, "Terms");
                OrderReference = (string)GetSQLValue(rdrOrder, "RefNum");
                Comments = (string)GetSQLValue(rdrOrder, "Comments");
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="rdrItem"></param>
            public void LoadLineItemsFromSQL(SqlDataReader rdrItem, bool treatAsNew)
            {
                LineItemsClass d = new LineItemsClass
                {
                    // If treatAsNew is set then we need to treat this as a new record rather than a read of an existing record.
                    //   This is used when creating a copy of an existing item for a new PO
                    OrderItemID = (treatAsNew ? 0 : (int)GetSQLValue(rdrItem, "RecID")),
                    Quantity = (int)GetSQLValue(rdrItem, "Qty"),
                    Units = (string)GetSQLValue(rdrItem, "Units"),
                    Description = (string)GetSQLValue(rdrItem, "Description"),
                    UnitPrice = (double)GetSQLValue(rdrItem, "UnitPrice"),
                    Purpose = (string)GetSQLValue(rdrItem, "Purpose"),
                    ChargeTo = (string)GetSQLValue(rdrItem, "ChargeTo"),
                    Classification = (string)GetSQLValue(rdrItem, "Class"),
                    IsTaxable = ((string)GetSQLValue(rdrItem, "Taxable") == "1" ? true : false) // This is a quirk of the OrderDetails schema which has field Taxable as a VARCHAR
                };
                d.TotalPrice = d.Quantity * d.UnitPrice;
                LineItems.Add(d);
                ComputeOrderTotal();
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="row"></param>
            /// <param name="d"></param>
            internal void UpdateRow(int row, LineItemsClass d)
            {
                // Update an order detail record with new data (d).
                LineItems[row].UpdateRow(d);
                OrderIsSaved = false;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="row"></param>
            internal void RemoveRow(int row)
            {
                try
                {
                    LineItems.RemoveAt(row);
                }
                catch
                {
                    // Ignore any exceptions as these will occur on grid clear and other initializations.
                }
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="row"></param>
            /// <returns></returns>
            internal LineItemsClass GetLineItems(int row)
            {
                return (LineItems[row]);
            }

            /// <summary>
            /// 
            /// </summary>
            /// <returns></returns>
            internal LineItemsClass AddDetailRecord()
            {
                OrderIsSaved = false;
                LineItems.Add(new LineItemsClass());
                return (LineItems[LineItems.Count - 1]);
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="pONum"></param>
            internal void UpdatePONumber(int pONum)
            {
                PONumber = pONum;
            }

            /// <summary>
            /// 
            /// </summary>
            public int NumLineItems
            {
                get { return LineItems.Count; }
            }

        }
        #endregion

        #region OrderItemClass
        /// <summary>
        /// 
        /// </summary>
        private class OrderItemClass
        {
            public int SelectedRow { get; set; }
            public int SelectedCol { get; set; }
            public int SelectedOrderID { get; set; }
            public int SelectItemRecordID { get; set; }
            public bool SingleItemOnly { get; set; }

            public OrderItemClass(int row, int col)
            {
                SelectedRow = row;
                SelectedCol = col;
            }
        }
        #endregion

        #region LineItemsClass

        /// <summary>
        /// 
        /// </summary>
        private class LineItemsClass
        {
            public int OrderItemID { get; set; }
            public int Quantity { get; set; }
            public string Units { get; set; }
            public string Description { get; set; }
            public double UnitPrice { get; set; }
            public double TotalPrice { get; set; }
            public bool IsTaxable { get; set; }
            public string ChargeTo { get; set; }
            public string Classification { get; set; }
            public string Purpose { get; set; }

            public LineItemsClass()
            {
                OrderItemID = 0;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="d"></param>
            internal void UpdateRow(LineItemsClass d)
            {
                OrderItemID = d.OrderItemID;
                Quantity = d.Quantity;
                Units = d.Units;
                Description = d.Description;
                UnitPrice = d.UnitPrice;
                TotalPrice = (double)Quantity * UnitPrice;
                IsTaxable = d.IsTaxable;
                ChargeTo = d.ChargeTo;
                Classification = d.Classification;
                Purpose = d.Purpose;
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


    }
}