using System;

namespace PurchaseOrders
{
    partial class frmMain
    {
        // ALL spreadsheet constructors are put here so that we don't clutter the main code with too much stuff.
        private clsSpreadsheet CreatePurchaseOrderSpreadsheet()
        {
            // Create a nice-looking Excel spreadsheet that can be used as a printable Purchase Order
            // NOTE that the complexity of this spreadsheet doesn't render itself to very much optimization.
            //    SOME optimization can be/has been implemented, using .Range functionality as much as possible.
            //    Additional optimizations might be possible by declaring a data array of objects, 
            //    saving data to the objects and writing the data as a single range rather than saving to individual cells.

            LblCreatingSpreadsheet.Visible = true;
            clsSpreadsheet xlPO = new clsSpreadsheet();

            // Landscape mode

            xlPO.OrientationLandscape();
            xlPO.MarginsNormal();
            xlPO.FitToPagesWide(1);

            // Expand Row 2 to 45.0, column K to 10.71, and column L to 13.57
            xlPO.SetRowHeight(2, 45);
            xlPO.SetColumnWidth(11, 10.71);
            xlPO.SetColumnWidth(12, 13.57);
            xlPO.FormatRangeFont(1, 1, 100, 25, "Calibri");

            // Load Images
            // TBD Try embedding these as resources and loading them internally.  That removes
            //    any dependencies on where they exist in the program.
            string f = AppDomain.CurrentDomain.BaseDirectory + "BufNewsLogo.jpg";
            xlPO.DisplayImage(f, 0, 0, 188, 217);
            f = AppDomain.CurrentDomain.BaseDirectory + "ForOfficeUseOnly.jpg";
            xlPO.DisplayImage(f, 690, 0, 314, 228); // TBD

            // Purchase Order #:  Merge K2/K3 and Bold the PO cell (K2)

            Point POLABEL = new Point(6, 2);
            xlPO.SetCellValue(POLABEL.Y, POLABEL.X, "PURCHASE ORDER");
            xlPO.FormatCellFontSize(POLABEL.Y, POLABEL.X, 28);
            xlPO.FormatCellAlignLeft(POLABEL.Y, POLABEL.X);
            xlPO.FormatCellAlignMiddle(POLABEL.Y, POLABEL.X);

            Point POCELL = new Point(11, 2);
            xlPO.MergeCells(POCELL.Y, POCELL.X, POCELL.Y, POCELL.X + 1);
            xlPO.FormatRangeBox(POCELL.Y, POCELL.X, POCELL.Y, POCELL.X + 1, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Color.Black);
            xlPO.FormatCellFontSize(POCELL.Y, POCELL.X, 36);
            xlPO.FormatCellBold(POCELL.Y, POCELL.X);
            xlPO.FormatCellAlignCenter(POCELL.Y, POCELL.X);
            xlPO.FormatCellAlignMiddle(POCELL.Y, POCELL.X);

            // Vendor Box

            Point VENDORBOXUL = new Point(6, 5);
            Point VENDORBOXLR = new Point(12, 13);
            xlPO.FormatRangeBox(VENDORBOXUL.Y, VENDORBOXUL.X, VENDORBOXLR.Y, VENDORBOXLR.X, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick, Color.Black);
            xlPO.FormatRangeBold(VENDORBOXUL.Y, VENDORBOXUL.X, VENDORBOXLR.Y, VENDORBOXLR.X);
            xlPO.FormatRangeItalic(VENDORBOXUL.Y, VENDORBOXUL.X, VENDORBOXLR.Y, VENDORBOXLR.X);
            xlPO.SetCellValue(VENDORBOXUL.Y - 1, VENDORBOXUL.X, "Vendor");

            // DataGrid Headings

            const int COLQUANTITY = 1;
            const int COLUNITDESC = 2;
            const int COLRECEIVED = 3;
            const int COLDESCSTRT = 4;
            const int COLDESCREND = 10;
            const int COLUNITPRIC = 11;
            const int COLTOTLPRIC = 12;
            const int COLTAXABLEX = 13;
            const int COLCHGTITLE = 14;
            const int COLPURTITLE = 14;
            const int COLCHPUSTRT = 15;
            const int COLCHARGEND = 16;
            const int COLCLASSTIT = 17;
            const int COLCLASSTRT = 18;
            const int COLCLASSEND = 20;
            const int COLPURPSEND = 20;

            const int ROWDGHEADNG = 15;

            // WARNING:  ALIGNMENT MUST COME FIRST
            //  Tryin to align after merge will cause additional merging to take place.
            // ALWAYS ALIGN FIRST!!!!!
            xlPO.FormatRangeBackColor(ROWDGHEADNG, COLQUANTITY, ROWDGHEADNG, COLCLASSEND, Color.SlateGray);
            xlPO.FormatRangeFontColor(ROWDGHEADNG, COLQUANTITY, ROWDGHEADNG, COLCLASSEND, Color.White);
            xlPO.FormatRangeAlignCenter(ROWDGHEADNG, COLQUANTITY, ROWDGHEADNG, COLCLASSEND);
            xlPO.FormatRangeBold(ROWDGHEADNG, COLQUANTITY, ROWDGHEADNG, COLCLASSEND);
            xlPO.MergeCells(ROWDGHEADNG, COLDESCSTRT, ROWDGHEADNG, COLDESCREND);
            xlPO.MergeCells(ROWDGHEADNG, COLCHPUSTRT, ROWDGHEADNG, COLCLASSEND);

            xlPO.FormatCellBox(ROWDGHEADNG, COLQUANTITY, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatCellBox(ROWDGHEADNG, COLUNITDESC, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatCellBox(ROWDGHEADNG, COLRECEIVED, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatRangeBox(ROWDGHEADNG, COLDESCSTRT, ROWDGHEADNG, 10, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatCellBox(ROWDGHEADNG, COLUNITPRIC, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatCellBox(ROWDGHEADNG, COLTOTLPRIC, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatCellBox(ROWDGHEADNG, COLTAXABLEX, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatCellBox(ROWDGHEADNG, COLCHGTITLE, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);
            xlPO.FormatRangeBox(ROWDGHEADNG, COLCHPUSTRT, ROWDGHEADNG, 20, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);

            xlPO.SetCellValue(ROWDGHEADNG, COLQUANTITY, "Qty");
            xlPO.SetCellValue(ROWDGHEADNG, COLUNITDESC, "Units");
            xlPO.SetCellValue(ROWDGHEADNG, COLRECEIVED, "#Rec'd");
            xlPO.SetCellValue(ROWDGHEADNG, COLDESCSTRT, "Description");
            xlPO.SetCellValue(ROWDGHEADNG, COLUNITPRIC, "Unit Price");
            xlPO.SetCellValue(ROWDGHEADNG, COLTOTLPRIC, "Total Price");
            xlPO.SetCellValue(ROWDGHEADNG, COLTAXABLEX, "Taxable");
            xlPO.SetCellValue(ROWDGHEADNG, COLCHPUSTRT, "Accounting Info");

            // Format the DataGrid rows.  Display a minimum of 8 rows whether or not they are all populated.

            int ROWGRIDSTART = ROWDGHEADNG + 1;
            int NumRows = ((GrdOrderDetails.RowCount - 1 <= 8) ? 8 : GrdOrderDetails.RowCount - 1);

            // Center some cells
            int lastrow = ROWGRIDSTART + 2 * (NumRows - 1) + 1;
            xlPO.FormatRangeAlignMiddle(ROWGRIDSTART, COLQUANTITY, lastrow, COLCLASSEND);         // Start with all rows vertically centered
            xlPO.FormatRangeAlignCenter(ROWGRIDSTART, COLQUANTITY, lastrow, COLUNITDESC);         // Qty and Units
            xlPO.FormatRangeAlignCenter(ROWGRIDSTART, COLTAXABLEX, lastrow, COLCHGTITLE);         // Taxable and the Charge/Purpose headings
            xlPO.FormatRangeAlignCenter(ROWGRIDSTART, COLCLASSTIT, lastrow, COLCLASSTIT);         // Classification
                                                                                                  // Left align some cells
            xlPO.FormatRangeAlignLeft(ROWGRIDSTART, COLDESCSTRT, lastrow, COLDESCSTRT);           // Description
            xlPO.FormatRangeAlignLeft(ROWGRIDSTART, COLCHPUSTRT, lastrow, COLCHPUSTRT);           // Charge and Purpose
            xlPO.FormatRangeAlignLeft(ROWGRIDSTART, COLCLASSTRT, lastrow, COLCLASSTRT);           // Classification
                                                                                                  // Right align some cells
            xlPO.FormatRangeAlignRight(ROWGRIDSTART, COLUNITPRIC, lastrow, COLTOTLPRIC);          // Unit Price and Total Price

            // Format some cells with currency format
            xlPO.FormatRangeCurrency(ROWGRIDSTART, COLUNITPRIC, lastrow, COLTOTLPRIC);

            // Set background color on some cells
            xlPO.FormatRangeBackColor(ROWGRIDSTART, COLCHGTITLE, lastrow, COLCHGTITLE, Color.SlateGray);
            xlPO.FormatRangeBackColor(ROWGRIDSTART, COLCLASSTIT, lastrow, COLCLASSTIT, Color.SlateGray);
            // Set font color on these same cells
            xlPO.FormatRangeFontColor(ROWGRIDSTART, COLCHGTITLE, lastrow, COLCHGTITLE, Color.White);
            xlPO.FormatRangeFontColor(ROWGRIDSTART, COLCLASSTIT, lastrow, COLCLASSTIT, Color.White);
            // Set font bold on these same cells
            xlPO.FormatRangeBold(ROWGRIDSTART, COLCHGTITLE, lastrow, COLCHGTITLE);
            xlPO.FormatRangeBold(ROWGRIDSTART, COLCLASSTIT, lastrow, COLCLASSTIT);
            // Set font italic on these same cells
            xlPO.FormatRangeItalic(ROWGRIDSTART, COLCHGTITLE, lastrow, COLCHGTITLE);
            xlPO.FormatRangeItalic(ROWGRIDSTART, COLCLASSTIT, lastrow, COLCLASSTIT);

            // Apply word wrap to the Description column

            xlPO.FormatRangeWrap(ROWGRIDSTART, COLDESCSTRT, lastrow, COLDESCSTRT);

            for (int i = 0; i < NumRows; i++)
            {
                int row1 = ROWGRIDSTART + 2 * i;
                int row2 = row1 + 1;

                // If columns have already been merged across a set of rows, then merging adjacent rows together causes the already-merged cells to merge
                //   into a single merged block of cells.

                // WARNING:  ALIGNMENT MUST COME FIRST
                //  Tryin to align after merge will cause additional merging to take place.
                // ALWAYS ALIGN FIRST!!!!!

                // Center some cells
                xlPO.FormatRangeAlignCenter(row1, COLTAXABLEX, row1, COLTAXABLEX);
                xlPO.FormatRangeAlignCenter(row1, COLCHGTITLE, row2, COLCHGTITLE);
                xlPO.FormatRangeAlignCenter(row1, COLCLASSTIT, row1, COLCLASSTIT);
                // Left align some cells
                xlPO.FormatRangeAlignLeft(row1, COLDESCSTRT, row1, COLDESCSTRT);
                xlPO.FormatRangeAlignLeft(row1, COLCHPUSTRT, row1, COLCHPUSTRT);
                xlPO.FormatRangeAlignLeft(row2, COLCHPUSTRT, row2, COLCHPUSTRT);
                xlPO.FormatRangeAlignLeft(row1, COLCLASSTRT, row1, COLCLASSTRT);
                // Right align some cells
                xlPO.FormatRangeAlignRight(row1, COLUNITPRIC, row1, COLTOTLPRIC);

                xlPO.MergeCells(row1, COLQUANTITY, row2, COLQUANTITY);
                xlPO.MergeCells(row1, COLUNITDESC, row2, COLUNITDESC);
                xlPO.MergeCells(row1, COLRECEIVED, row2, COLRECEIVED);
                xlPO.MergeCells(row1, COLDESCSTRT, row2, COLDESCREND);      // Description
                xlPO.MergeCells(row1, COLUNITPRIC, row2, COLUNITPRIC);
                xlPO.MergeCells(row1, COLTOTLPRIC, row2, COLTOTLPRIC);
                xlPO.MergeCells(row1, COLTAXABLEX, row2, COLTAXABLEX);
                xlPO.MergeCells(row1, COLCHPUSTRT, row1, COLCHARGEND);      // Charge
                xlPO.MergeCells(row2, COLCHPUSTRT, row2, COLPURPSEND);      // Purpose
                xlPO.MergeCells(row1, COLCLASSTRT, row1, COLCLASSEND);      // Class

                // Format some cells with currency format
                xlPO.FormatRangeCurrency(row1, COLUNITPRIC, row2, COLTOTLPRIC);

                // Set background color on some cells
                xlPO.FormatRangeBackColor(row1, COLCHGTITLE, row2, COLCHGTITLE, Color.SlateGray);
                xlPO.FormatRangeBackColor(row1, COLCLASSTIT, row1, COLCLASSTIT, Color.SlateGray);
                // Set font color on these same cells
                xlPO.FormatRangeFontColor(row1, COLCHGTITLE, row2, COLCHGTITLE, Color.White);
                xlPO.FormatRangeFontColor(row1, COLCLASSTIT, row1, COLCLASSTIT, Color.White);
                // Set font bold on these same cells
                xlPO.FormatRangeBold(row1, COLCHGTITLE, row2, COLCHGTITLE);
                xlPO.FormatRangeBold(row1, COLCLASSTIT, row1, COLCLASSTIT);
                // Set font italic on these same cells
                xlPO.FormatRangeItalic(row1, COLCHGTITLE, row2, COLCHGTITLE);
                xlPO.FormatRangeItalic(row1, COLCLASSTIT, row1, COLCLASSTIT);
                // Populate these cells with text
                xlPO.SetCellValue(row1, COLCHGTITLE, "Charge");
                xlPO.SetCellValue(row2, COLPURTITLE, "Purpose");
                xlPO.SetCellValue(row1, COLCLASSTIT, "Class");

                // Add Borders
                xlPO.FormatRangeBox(row1, COLQUANTITY, row2, COLQUANTITY, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <quantity>
                xlPO.FormatRangeBox(row1, COLUNITDESC, row2, COLUNITDESC, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <units>
                xlPO.FormatRangeBox(row1, COLRECEIVED, row2, COLRECEIVED, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <received>
                xlPO.FormatRangeBox(row1, COLDESCSTRT, row2, COLDESCREND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <description>
                xlPO.FormatRangeBox(row1, COLUNITPRIC, row2, COLUNITPRIC, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <unit price>
                xlPO.FormatRangeBox(row1, COLTOTLPRIC, row2, COLTOTLPRIC, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <total price>
                xlPO.FormatRangeBox(row1, COLTAXABLEX, row2, COLTAXABLEX, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <taxable>
                xlPO.FormatRangeBox(row1, COLCHGTITLE, row1, COLCHGTITLE, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // Charge Title
                xlPO.FormatRangeBox(row1, COLCHPUSTRT, row1, COLCHARGEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <charge>
                xlPO.FormatRangeBox(row1, COLCLASSTIT, row1, COLCLASSTIT, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // Class Title
                xlPO.FormatRangeBox(row1, COLCLASSTRT, row1, COLCLASSEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <classification>
                xlPO.FormatRangeBox(row2, COLPURTITLE, row2, COLPURTITLE, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // Purpose Title
                xlPO.FormatRangeBox(row2, COLCHPUSTRT, row2, COLPURPSEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <purpose>

                // Populate these cells with text
                xlPO.SetCellValue(row1, COLCHGTITLE, "Charge");
                xlPO.SetCellValue(row2, COLPURTITLE, "Purpose");
                xlPO.SetCellValue(row1, COLCLASSTIT, "Class");
            }

            // Totals Row

            int ROWGRIDTOTALS = ROWGRIDSTART + 2 * NumRows;
            xlPO.FormatRangeBackColor(ROWGRIDTOTALS, COLQUANTITY, ROWGRIDTOTALS, COLCLASSEND, Color.SlateGray);
            xlPO.FormatRangeFontColor(ROWGRIDTOTALS, COLQUANTITY, ROWGRIDTOTALS, COLCLASSEND, Color.White);
            xlPO.FormatCellBold(ROWGRIDTOTALS, COLUNITPRIC);
            xlPO.FormatCellItalic(ROWGRIDTOTALS, COLUNITPRIC);
            xlPO.SetCellValue(ROWGRIDTOTALS, COLUNITPRIC, "Order Total");

            // Add all contract boxes at the bottom of the order

            //Merge Rows together
            int ROW1SIGS = ROWGRIDTOTALS + 2;
            int ROW2SIGS = ROW1SIGS + 2;
            int ROW3SIGS = ROW1SIGS + 4;

            const int COLORDEREDBY = 1;
            const int COLDELIVERTO = 1;
            const int COLAPPROVALX = 1;
            const int COLORDBYSTRT = 2;
            const int COLDELTOSTRT = 2;
            const int COLORDRBYEND = 7;
            const int COLDELTOSEND = 7;
            const int COLAPPROVEND = 7;
            const int COLEXTENSION = 8;
            const int COLDEPARTMNT = 8;
            const int COLRECEIVDBY = 8;
            const int COLEXTENSTRT = 9;
            const int COLDEPTMSTRT = 9;
            const int COLEXTENSEND = 11;
            const int COLDEPARTEND = 11;
            const int COLRCVDBYEND = 11;
            const int COLORDERDATE = 12;
            const int COLRECVDDATE = 12;
            const int COLTRMTITSTR = 14;
            const int COLREFTITSTR = 14;
            const int COLCOMTITSTR = 14;
            const int COLTRMTITEND = 15;
            const int COLREFTITEND = 15;
            const int COLCOMTITEND = 15;
            const int COLTERMSSTRT = 16;
            const int COLREFNUSTRT = 16;
            const int COLCOMMNSTRT = 16;
            const int COLTERMSXEND = 20;
            const int COLREFNUMEND = 20;
            const int COLCOMMNTEND = 20;

            const int FONTSMALL = 9;
            const int FONTLARGE = 16;

            xlPO.MergeCells(ROW1SIGS, COLORDEREDBY, ROW1SIGS + 1, COLORDEREDBY);  // Ordered By
            xlPO.MergeCells(ROW1SIGS, COLORDBYSTRT, ROW1SIGS + 1, COLORDRBYEND);  // <ordered by>
            xlPO.MergeCells(ROW1SIGS, COLEXTENSION, ROW1SIGS + 1, COLEXTENSION);  // Extension
            xlPO.MergeCells(ROW1SIGS, COLEXTENSTRT, ROW1SIGS + 1, COLEXTENSEND);  // <extension #>
            xlPO.MergeCells(ROW1SIGS, COLTRMTITSTR, ROW1SIGS + 1, COLTRMTITEND);  // Terms
            xlPO.MergeCells(ROW1SIGS, COLTERMSSTRT, ROW1SIGS + 1, COLTERMSXEND);  // <terms>
            xlPO.MergeCells(ROW2SIGS, COLDELIVERTO, ROW2SIGS + 1, COLDELIVERTO);  // Deliver To
            xlPO.MergeCells(ROW2SIGS, COLDELTOSTRT, ROW2SIGS + 1, COLDELTOSEND);  // <deliver to name>
            xlPO.MergeCells(ROW2SIGS, COLDEPARTMNT, ROW2SIGS + 1, COLDEPARTMNT);  // Dept
            xlPO.MergeCells(ROW2SIGS, COLDEPTMSTRT, ROW2SIGS + 1, COLDEPARTEND);  // <dept>
            xlPO.MergeCells(ROW2SIGS, COLREFTITSTR, ROW2SIGS + 1, COLREFTITEND);  // Ref #
            xlPO.MergeCells(ROW2SIGS, COLREFNUSTRT, ROW2SIGS + 1, COLREFNUMEND);  // <ref #>
            xlPO.MergeCells(ROW3SIGS, COLAPPROVALX, ROW3SIGS + 1, COLAPPROVEND);  // Approval
            xlPO.MergeCells(ROW3SIGS, COLRECEIVDBY, ROW3SIGS + 1, COLRCVDBYEND);  // Received By
            xlPO.MergeCells(ROW3SIGS, COLORDERDATE, ROW3SIGS + 1, COLORDERDATE);  // Date
            xlPO.MergeCells(ROW3SIGS, COLCOMTITSTR, ROW3SIGS + 1, COLCOMTITEND);  // Comments
            xlPO.MergeCells(ROW3SIGS, COLCOMMNSTRT, ROW3SIGS + 1, COLCOMMNTEND);  // <comments>

            // Add Boxes 
            xlPO.FormatRangeBox(ROW1SIGS, COLORDEREDBY, ROW1SIGS + 1, COLORDRBYEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // <ordered by>
            xlPO.FormatRangeBox(ROW1SIGS, COLEXTENSION, ROW1SIGS + 1, COLEXTENSEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // Extension
            xlPO.FormatRangeBox(ROW1SIGS, COLORDERDATE, ROW1SIGS + 1, COLORDERDATE, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // Order Date    
            xlPO.FormatRangeBox(ROW1SIGS, COLTRMTITSTR, ROW1SIGS + 1, COLTRMTITEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // Terms
            xlPO.FormatRangeBox(ROW1SIGS, COLTERMSSTRT, ROW1SIGS + 1, COLTERMSXEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // <terms>
            xlPO.FormatRangeBox(ROW2SIGS, COLDELIVERTO, ROW2SIGS + 1, COLDELTOSEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // Deliver To
            xlPO.FormatRangeBox(ROW2SIGS, COLDEPARTMNT, ROW2SIGS + 1, COLDEPARTEND + 1, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);  // <dept>
            xlPO.FormatRangeBox(ROW2SIGS, COLREFTITSTR, ROW2SIGS + 1, COLREFTITEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // Ref #
            xlPO.FormatRangeBox(ROW2SIGS, COLREFNUSTRT, ROW2SIGS + 1, COLREFNUMEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // <ref #>
            xlPO.FormatRangeBox(ROW3SIGS, COLAPPROVALX, ROW3SIGS + 1, COLAPPROVEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // <approval>
            xlPO.FormatRangeBox(ROW3SIGS, COLRECEIVDBY, ROW3SIGS + 1, COLRCVDBYEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // <Received By>
            xlPO.FormatRangeBox(ROW3SIGS, COLRECVDDATE, ROW3SIGS + 1, COLRECVDDATE, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // <date>
            xlPO.FormatRangeBox(ROW3SIGS, COLCOMTITSTR, ROW3SIGS + 1, COLCOMTITEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // Comments
            xlPO.FormatRangeBox(ROW3SIGS, COLCOMMNSTRT, ROW3SIGS + 1, COLCOMMNTEND, Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Color.Black);      // <comments>

            // Shade a few boxes
            xlPO.FormatRangeBackColor(ROW1SIGS, COLTRMTITSTR, ROW1SIGS + 1, COLTRMTITEND, Color.SlateGray);         // Terms
            xlPO.FormatRangeBackColor(ROW2SIGS, COLREFTITSTR, ROW2SIGS + 1, COLREFTITEND, Color.SlateGray);         // Ref #              
            xlPO.FormatRangeBackColor(ROW3SIGS, COLCOMTITSTR, ROW3SIGS + 1, COLCOMTITEND, Color.SlateGray);         // Comments

            xlPO.FormatRangeFontColor(ROW1SIGS, COLTRMTITSTR, ROW1SIGS + 1, COLTRMTITEND, Color.White);             // Terms
            xlPO.FormatRangeFontColor(ROW2SIGS, COLREFTITSTR, ROW2SIGS + 1, COLREFTITEND, Color.White);             // Ref #
            xlPO.FormatRangeFontColor(ROW3SIGS, COLCOMTITSTR, ROW3SIGS + 1, COLCOMTITEND, Color.White);             // Comments

            // Bold a few boxes
            xlPO.FormatCellBold(ROW1SIGS, COLORDBYSTRT);        // <ordered by>
            xlPO.FormatCellBold(ROW1SIGS, COLEXTENSTRT);        // <extension>
            xlPO.FormatCellBold(ROW1SIGS, COLTRMTITSTR);        // Terms
            xlPO.FormatCellBold(ROW1SIGS, COLTERMSSTRT);        // <terms>
            xlPO.FormatCellBold(ROW1SIGS + 1, COLORDERDATE);    // <order date>
            xlPO.FormatCellBold(ROW2SIGS, COLDELTOSTRT);        // <deliver to
            xlPO.FormatCellBold(ROW2SIGS, COLDEPTMSTRT);        // <dept>
            xlPO.FormatCellBold(ROW2SIGS, COLREFTITSTR);        // Ref #
            xlPO.FormatCellBold(ROW2SIGS, COLREFNUSTRT);        // <ref #>
            xlPO.FormatCellBold(ROW3SIGS, COLCOMTITSTR);        // Comments

            // Italicize a few boxes (then undo italics for the order date)
            xlPO.FormatRangeItalic(ROW1SIGS, COLORDEREDBY, ROW3SIGS + 1, COLRECVDDATE);
            xlPO.FormatRangeItalic(ROW1SIGS, COLREFNUSTRT, ROW2SIGS + 1, COLCOMMNTEND);
            xlPO.FormatCellNoItalic(ROW1SIGS + 1, COLORDERDATE);

            // Set a few font sizes
            xlPO.FormatRangeFontSize(ROW1SIGS, COLORDEREDBY, ROW3SIGS + 1, COLORDEREDBY, FONTSMALL);
            xlPO.FormatRangeFontSize(ROW1SIGS, COLORDBYSTRT, ROW2SIGS + 1, COLORDBYSTRT, FONTLARGE);
            xlPO.FormatRangeFontSize(ROW1SIGS, COLEXTENSION, ROW3SIGS + 1, COLEXTENSION, FONTSMALL);
            xlPO.FormatRangeFontSize(ROW1SIGS, COLEXTENSTRT, ROW2SIGS + 1, COLEXTENSTRT, FONTLARGE);
            xlPO.FormatRangeFontSize(ROW1SIGS, COLTERMSSTRT, ROW2SIGS + 1, COLTERMSSTRT, FONTLARGE);
            xlPO.FormatCellFontSize(ROW1SIGS, COLORDERDATE, FONTSMALL);     // Order Date
            xlPO.FormatCellFontSize(ROW3SIGS, COLRECVDDATE, FONTSMALL);     // Rec'd Date

            // Set up appropriate alignment
            xlPO.FormatRangeAlignTop(ROW1SIGS, COLORDEREDBY, ROW3SIGS + 1, COLRECVDDATE);
            xlPO.FormatRangeAlignMiddle(ROW1SIGS, COLTRMTITSTR, ROW3SIGS + 1, COLCOMMNTEND);
            xlPO.FormatCellAlignMiddle(ROW1SIGS, COLORDBYSTRT);         // <ordered by>
            xlPO.FormatCellAlignMiddle(ROW1SIGS, COLEXTENSTRT);         // <extension>
            xlPO.FormatCellAlignMiddle(ROW1SIGS + 1, COLORDERDATE);     // <order date>
            xlPO.FormatCellAlignMiddle(ROW2SIGS, COLDELTOSTRT);         // <deliver to>
            xlPO.FormatCellAlignMiddle(ROW2SIGS, COLDEPTMSTRT);         // <dept>

            xlPO.FormatRangeAlignLeft(ROW1SIGS, COLORDEREDBY, ROW3SIGS + 1, COLRECVDDATE);
            xlPO.FormatRangeAlignLeft(ROW1SIGS, COLTERMSSTRT, ROW3SIGS + 1, COLCOMMNTEND);
            xlPO.FormatRangeAlignCenter(ROW1SIGS, COLTRMTITSTR, ROW3SIGS + 1, COLTRMTITEND);
            xlPO.FormatCellAlignCenter(ROW1SIGS, COLORDBYSTRT);         // <ordered by>
            xlPO.FormatCellAlignCenter(ROW1SIGS, COLEXTENSTRT);         // <extension>
            xlPO.FormatCellAlignCenter(ROW1SIGS + 1, COLORDERDATE);     // <order date>
            xlPO.FormatCellAlignCenter(ROW2SIGS, COLDELTOSTRT);         // <deliver to>
            xlPO.FormatCellAlignCenter(ROW2SIGS, COLDEPTMSTRT);         // <dept>

            // Wrap the comments
            xlPO.FormatCellText(ROW3SIGS, COLCOMMNSTRT);
            xlPO.FormatCellWrap(ROW3SIGS, COLCOMMNSTRT);                // <comments>

            // Add constant terms
            xlPO.SetCellValue(ROW1SIGS, COLORDEREDBY, "Ordered By");
            xlPO.SetCellValue(ROW1SIGS, COLEXTENSION, "Extension");
            xlPO.SetCellValue(ROW1SIGS, COLORDERDATE, "Order Date");
            xlPO.SetCellValue(ROW1SIGS, COLTRMTITSTR, "Terms");
            xlPO.SetCellValue(ROW2SIGS, COLDELIVERTO, "Deliver To");
            xlPO.SetCellValue(ROW2SIGS, COLDEPARTMNT, "Dept");
            xlPO.SetCellValue(ROW2SIGS, COLREFTITSTR, "Ref #");
            xlPO.SetCellValue(ROW3SIGS, COLAPPROVALX, "Approval");
            xlPO.SetCellValue(ROW3SIGS, COLRECEIVDBY, "Received By");
            xlPO.SetCellValue(ROW3SIGS, COLRECVDDATE, "Date");
            xlPO.SetCellValue(ROW3SIGS, COLCOMTITSTR, "Comments");

            // POPULATE ALL PO DATA

            // Populate the PO number
            OrderClass o = CurrentOrder;
            xlPO.SetCellValue(POCELL.Y, POCELL.X, "D" + o.PONumber.Value.ToString("D5"));

            VendorClass v = CurrentVendors.VendorList[CurrentVendors.SelectedListIndex];
            // Populate the Vendor info
            xlPO.SetCellValue(VENDORBOXUL.Y + 1, VENDORBOXUL.X + 1, v.VendorName);
            xlPO.SetCellValue(VENDORBOXUL.Y + 2, VENDORBOXUL.X + 1, v.AddrLine1);
            int nextrow = 3;
            if (v.AddrLine2 != null)
            {
                if (v.AddrLine2.Length > 0)
                {
                    xlPO.SetCellValue(VENDORBOXUL.Y + 3, VENDORBOXUL.X + 1, v.AddrLine2);
                    nextrow = 4;
                }
            }
            xlPO.SetCellValue(VENDORBOXUL.Y + nextrow, VENDORBOXUL.X + 1, v.City + ", " + v.State + "  " + v.Zip);
            if (v.Contact != null)
            {
                if (v.Contact.Length > 0)
                {
                    xlPO.SetCellValue(VENDORBOXUL.Y + nextrow + 2, VENDORBOXUL.X + 1, "Attn: " + v.Contact);
                    xlPO.SetCellValue(VENDORBOXUL.Y + nextrow + 3, VENDORBOXUL.X + 1, v.Phone);
                }
            }
            if (v.AcctNum != null)
            {
                if (v.AcctNum.Length > 0)
                {
                    xlPO.SetCellValue(VENDORBOXUL.Y + 8, VENDORBOXUL.X + 2, "News Acct # " + v.AcctNum);
                }
            }

            // Populate spreadsheet rows from the PO data

            for (int i = 0; i < o.NumLineItems; i++)
            {
                int row1start = ROWGRIDSTART + 2 * i;
                int row2start = row1start + 1;

                LineItemsClass d = o.GetLineItems(i);
                xlPO.SetCellValue(row1start, COLQUANTITY, d.Quantity);
                xlPO.SetCellValue(row1start, COLUNITDESC, d.Units);
                xlPO.SetCellValue(row1start, COLDESCSTRT, d.Description);
                xlPO.SetCellValue(row1start, COLUNITPRIC, d.UnitPrice);
                xlPO.SetCellValue(row1start, COLTOTLPRIC, d.TotalPrice);
                xlPO.SetCellValue(row1start, COLTAXABLEX, (d.IsTaxable ? "Yes" : "No"));
                xlPO.SetCellValue(row1start, COLCHPUSTRT, d.ChargeTo);
                xlPO.SetCellValue(row1start, COLCLASSTRT, d.Classification);
                xlPO.SetCellValue(row2start, COLCHPUSTRT, d.Purpose);
            }

            //Populate Order Total
            xlPO.SetCellValue(ROWGRIDTOTALS, COLTOTLPRIC, TxtPOTotal.Text);

            // Display username in the Ordered By box. 
            xlPO.SetCellValue(ROW1SIGS, COLORDBYSTRT, UserInfo.ADUserList[0].FirstName + " " + UserInfo.ADUserList[0].LastName);
            // Populate the rest of the signature area with data from the main PO form
            xlPO.SetCellValue(ROW2SIGS, COLDELTOSTRT, TxtDeliverto.Text);
            xlPO.SetCellValue(ROW1SIGS, COLEXTENSTRT, TxtDeliverToPhone.Text);
            xlPO.SetCellValue(ROW2SIGS, COLDEPTMSTRT, TxtDepartment.Text);
            xlPO.SetCellValue(ROW1SIGS, COLTERMSSTRT, TxtTerms.Text);
            xlPO.SetCellValue(ROW2SIGS, COLREFNUSTRT, "'" + TxtOrderReference.Text); // Force this to be text
            xlPO.SetCellValue(ROW3SIGS, COLCOMMNSTRT, TxtComments.Text);
            if (TxtDate.TextLength > 0)
            {
                bool dateokay = DateTime.TryParse(TxtDate.Text, out DateTime dt);
                if (dateokay)
                {
                    xlPO.SetCellValue(ROW1SIGS + 1, COLORDERDATE, dt.ToShortDateString());
                }
                else
                {
                    xlPO.SetCellValue(ROW1SIGS + 1, COLORDERDATE, TxtDate.Text);
                }
            }

            LblCreatingSpreadsheet.Visible = false;
            return (xlPO);
        }

    }
}