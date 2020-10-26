using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace BSGlobals
{
    // An attempt to produce a relatively generic spreadsheet class to encapsulate the following worksheet functions:
    //
    // FONT
    //    Size
    //    Bold
    //    Italic
    //    Backcolor
    //    Color
    // CELL
    //    Horizontal Justification
    //    Vertical Justification
    //    Format (text, number, currency, general)
    // MISC
    //    Wrap
    //    Merge
    //    Image
    //    Borders
    //    Column Width
    //    Row Height
    //    Fit to Page Width
    //    Margins (Narrow, Normal)

    // NOTES
    //    OLE process threads are notoriously difficult to consistently kill (they hang out in Task Manager as orphaned processes, taking up space).
    //       The Terminate() method appears to fix this by unilaterally killing the specific Excel process on termination.
    //    AutoFit rows and AutoFit columns do not appear to work when used according to Microsoft documentation.
    //    To keep this class as generic as possible, no log-writing is performed here.  Instead, most functions will return true normally,
    //       and false if any kind of exception occurred.  A public LastException string will be populated with the exception should one occur.
    //
    //    THIS IS NOT YET OPTIMIZED and as such it will run very slowly for any complex spreadsheet.  
    //       One optimization step that is sorely needed is a way to collect data flowing to individual cells and to encapsulate
    //       that data into a 2D array instead; and ONLY after the data has been fully assigned should the array be written to the spreadsheet via OLE.
    //       (Note that this will reduce OLE calls but not eliminate those related to formatting.  However, it will do wonders for large datasets that
    //       simply need to be written to Excel.

    public class Spreadsheet
    {

        #region Declarations

        // Enums types that allow us to remove any user dependencies on Microsoft.Office.Interop.Excel

        public enum LineStyle {
            Continuous = 1,
            Dash = -4115,
            DashDot = 4,
            DashDotDot = 5,
            Dot = -4118,
            Double = -4119,
            None = -4142,
            SlantDashDot = 13 };

        public enum BorderWeight {
            Hairline = 1,
            Medium = -4138,
            Thick = 4,
            Thin = 2 };

        public FontClass Font;
        public PageSetupClass PageSetup;
        public AlignmentClass Alignment;
        public FormatClass Format;
        public StylesClass Style;
        public FileClass File;
        public string LastException = "";

        private Excel.Application ExcelApp = new Excel.Application();
        private Excel.Workbook ExcelWorkbook;
        private Excel.Worksheet ExcelWorksheet;
        private Process ExcelProcess;

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        public int NumWorksheets { get; set; }
        #endregion

        #region Instantiation
        public Spreadsheet()
        {

            // Instantiate subclasses
            InstantiateCommon();

            // Add a workbook
            ExcelApp.Workbooks.Add();
            ExcelApp.DisplayAlerts = false;
            ExcelWorkbook = ExcelApp.Application.ActiveWorkbook;
            ExcelProcess = GetExcelProcess(ExcelApp);

            // Delete all but the first worksheet and point to it
            while (ExcelWorkbook.Worksheets.Count > 1)
            {
                ExcelWorksheet.Delete();
            }
            NumWorksheets = 1;
            ExcelWorksheet = ExcelWorkbook.Sheets[1];

            // Set standard row height and column width for all rows/columns on the spreadsheet

            ExcelWorksheet.Rows.UseStandardHeight = true;
            ExcelWorksheet.Columns.UseStandardWidth = true;
        }

        public Spreadsheet(string spreadsheetPath)
        {

            // Instantiate subclasses
            InstantiateCommon();

            // Get the workbook
            ExcelWorkbook =
                (Excel.Workbook)(ExcelApp.Workbooks._Open(spreadsheetPath, System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value)); // open the existing excel file
            //ExcelApp.Workbooks.Add();
            ExcelApp.DisplayAlerts = false;
            ExcelWorkbook = ExcelApp.Application.ActiveWorkbook;

            // Point to the first worksheet
            if (ExcelWorkbook.Worksheets.Count == 0)
            {
                ExcelWorkbook.Sheets.Add();
            }
            NumWorksheets = ExcelWorkbook.Worksheets.Count;
            ExcelWorksheet = ExcelWorkbook.Sheets[1];
            ExcelWorksheet.Activate();
        }

        private void InstantiateCommon()
        {
            // Instantiate subclasses

            Font = new FontClass(this);
            PageSetup = new PageSetupClass(this);
            Alignment = new AlignmentClass(this);
            Format = new FormatClass(this);
            Style = new StylesClass(this);
            File = new FileClass(this);

            // Create an Excel spreadsheet object
            if (ExcelApp == null)
            {
                MessageBox.Show("Excel is not properly installed.");
                return;
            }
        }

        private Process GetExcelProcess(Excel.Application excelApp)
        {
            GetWindowThreadProcessId(excelApp.Hwnd, out int id);
            return Process.GetProcessById(id);
        }

        #endregion

        #region Base Functions
        private Excel.Range SetRange(int startRow, int startCol, int endRow, int endCol)
        {
            Excel.Range range = ExcelWorksheet.Range[ExcelWorksheet.Cells[startRow, startCol], ExcelWorksheet.Cells[endRow, endCol]];
            range.Select();
            return (range);
        }

        /// <summary>
        /// Set a specific cell value
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool SetCellValue(int row, int col, object value)
        {
            try
            {
                ExcelWorksheet.Cells[row, col].Value = value;
                return (true);
            }
            catch 
            {
                return (false);
            }
        }

        /// <summary>
        /// Set the active sheet by sheet name
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public bool SetActiveSheet(string s)
        {
            try
            {
                ExcelWorksheet = ExcelWorkbook.Sheets[s];
                ExcelWorksheet.Activate();
                return (true);
            }
            catch
            {
                return (false);
            }
        }

        /// <summary>
        /// Set the active sheet by sheet index
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool SetActiveSheet(int index)
        {
            try
            {
                ExcelWorksheet = ExcelWorkbook.Sheets[index];
                ExcelWorksheet.Activate();
                return (true);
            }
            catch
            {
                return (false);
            }
        }

        /// <summary>
        ///  Display an image on the spreadsheet defined by the XY coordinates and size
        /// </summary>
        /// <param name="imageName"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns>true if no errors occurred.</returns>
        public bool InsertImage(string imageName, int x, int y, int width, int height)
        {
            try
            {
                ExcelWorksheet.Shapes.AddPicture(imageName, MsoTriState.msoFalse, MsoTriState.msoTrue, x, y, width, height);
                return (true);
            }
            catch (Exception ex)
            {
                LastException = ex.ToString();
                return (false);
            }
        }

        /// <summary>
        /// Display the spreadsheet on screen
        /// </summary>
        public void Show()
        {
            ExcelApp.PrintCommunication = true;
            ExcelApp.DisplayAlerts = true;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        /// <summary>
        /// Hide the spreadsheet
        /// </summary>
        public void Hide()
        {
            ExcelApp.Visible = false;
        }

        /// <summary>
        /// Print the spreadsheet to the default printer
        /// </summary>
        public void Print()
        {
            ExcelWorksheet.PrintOutEx();
        }

        /// <summary>
        /// Terminate the Excel process for this spreadsheet
        /// </summary>
        public void Terminate()
        {
            // Close the workbook (this always throws an exception)
            // Quit the excel app, then release any remaining COM objects
            //   and finally, kill the process (as it does not consistently die after releasing all objects)
            try { ExcelWorkbook.Close(); } catch { }
            try { ExcelApp.Quit(); } catch { }
            try
            {
                if (ExcelWorksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorksheet);
                if (ExcelWorkbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkbook);
                if (ExcelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                ExcelProcess.Kill();
            }
            catch
            {
                // TBD
            }
        }
        #endregion    

        #region FontClass
        public class FontClass
        {
            Spreadsheet SP;
            public FontClass(Spreadsheet sp)
            {
                SP = sp;
            }

            /// <summary>
            /// Set range of cells font name
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <param name="size"></param>
            /// <returns></returns>
            public bool Name(int startRow, int startCol, int endRow, int endCol, string fontName)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Font.Name = fontName;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell font name
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <param name="size"></param>
            /// <returns></returns>
            public bool Name(int row, int col, string fontName)
            {
                return (Name(row, col, row, col, fontName));
            }

            /// <summary>
            /// Set range of cells font size
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <param name="size"></param>
            /// <returns></returns>
            public bool Size(int startRow, int startCol, int endRow, int endCol, int size)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Font.Size = size;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell font size
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <param name="size"></param>
            /// <returns></returns>
            public bool Size(int row, int col, int size)
            {
                return (Size(row, col, row, col, size));
            }

            /// <summary>
            /// Set range of scells to font color
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <param name="color"></param>
            /// <returns></returns>
            public bool Color(int startRow, int startCol, int endRow, int endCol, Color color)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Font.Color = ColorTranslator.ToOle(color);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell font color
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <param name="color"></param>
            /// <returns></returns>
            public bool Color(int row, int col, Color color)
            {
                return (Color(row, col, row, col, color));
            }

            /// <summary>
            /// Set range of cells to bold
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool Bold(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Font.Bold = true;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell to bold
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Bold(int row, int col)
            {
                return (Bold(row, col, row, col));
            }

            /// <summary>
            /// Remove bold from range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool NoBold(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Font.Bold = false;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Remove bold from cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool NoBold(int row, int col)
            {
                return (NoBold(row, col, row, col));
            }

            /// <summary>
            /// Set range of cells to italic
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool Italic(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Font.Italic = true;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell text to italic
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Italic(int row, int col)
            {
                return (Italic(row, col, row, col));
            }

            /// <summary>
            /// Remove italics from range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool NoItalic(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Font.Italic = false;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Remove italics from cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool NoItalic(int row, int col)
            {
                return (NoItalic(row, col, row, col));
            }

        }
        #endregion

        #region FormatClass
        public class FormatClass
        {
            Spreadsheet SP;

            public FormatClass (Spreadsheet sp)
            {
                SP = sp;
            }

            /// <summary>
            /// Set cell format in the specified range to Currency ($0.00)
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool Currency(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Cells.Style = "Currency";
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell format to Currency ($0.00)
            /// 
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Currency(int row, int col)
            {
                return (Currency(row, col, row, col));
            }

            /// <summary>
            /// Set cell format in the specified range to Text
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool Text(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.NumberFormat = "@";
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell format to Text
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Text(int row, int col)
            {
                return (Text(row, col, row, col));
            }

            /// <summary>
            /// Set cell format in the specfied range to Number ("@")
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool Number(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.NumberFormat = "@";
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Set cell format to Number ("@")
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Number(int row, int col)
            {
                return (Number(row, col, row, col));
            }

            /// <summary>
            /// Formats range of cells as mm/dd/yyyy
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool ShortDate(int startRow, int startCol, int endRow, int endCol)
            {
                // MM/DD/YYYY only
                try
                {
                    Custom(startRow, startCol, endRow, endCol, "mm/dd/yyyy");
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Formats cell as mm/dd/yyyy
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool ShortDate(int row, int col)
            {
                return (ShortDate(row, col, row, col));
            }

            /// <summary>
            /// Formats range of cells as any valid Excel format.  Returns false if the formatting failed.
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <param name="dateFormat"></param>
            /// <returns></returns>
            /// 
            public bool Custom(int startRow, int startCol, int endRow, int endCol, string format)
            {                
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.NumberFormat = format;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Formats cell as any valid Excel format.  Returns false if the formatting failed.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Custom(int row, int col, string format)
            {
                return (Custom(row, col, row, col, format));
            }
        }
        #endregion

        #region AlignmentClass
        public class AlignmentClass
        {
            Spreadsheet SP;

            public AlignmentClass (Spreadsheet sp)
            {
                SP = sp;
            }
            /// <summary>
            /// Left-aligns range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <returns></returns>
            public bool Left(int startRow, int startCol, int endRow, int EndCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Left-aligns cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Left(int row, int col)
            {
                return (Left(row, col, row, col));
            }


            /// <summary>
            /// Center-aligns range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <returns></returns>
            public bool Center(int startRow, int startCol, int endRow, int EndCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Center-aligns cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Center(int row, int col)
            {
                return (Center(row, col, row, col));
            }


            /// <summary>
            /// Right-aligns range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <returns></returns>
            public bool Right(int startRow, int startCol, int endRow, int EndCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Right-aligns cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Right(int row, int col)
            {
                return (Right(row, col, row, col));
            }

            /// <summary>
            /// Top-aligns range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <returns></returns>
            public bool Top(int startRow, int startCol, int endRow, int EndCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Top-aligns cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Top(int row, int col)
            {
                return (Top(row, col, row, col));
            }


            /// <summary>
            /// (Vertical) Middle-aligns range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <returns></returns>
            public bool Middle(int startRow, int startCol, int endRow, int EndCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// (Vertical) Middle-aligns cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Middle(int row, int col)
            {
                return (Middle(row, col, row, col));
            }

            /// <summary>
            /// Bottom-aligns range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <returns></returns>
            public bool Bottom(int startRow, int startCol, int endRow, int EndCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Bottom-aligns cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Bottom(int row, int col)
            {
                return (Bottom(row, col, row, col));
            }

            /// <summary>
            /// Merges ranges of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <returns></returns>
            public bool MergeCells(int startRow, int startCol, int endRow, int EndCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.Merge();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Wraps text within range of cells
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <returns></returns>
            public bool Wrap(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.WrapText = true;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Wraps text in cell
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <returns></returns>
            public bool Wrap(int row, int col)
            {
                return (Wrap(row, col, row, col));
            }

        }
        #endregion

        #region StylesClass
        public class StylesClass
        {
            Spreadsheet SP;

            public StylesClass(Spreadsheet sp)
            {
                SP = sp;
            }

            /// <summary>
            /// Sets row height in points (72 points/inch)
            /// </summary>
            /// <param name="rowNum"></param>
            /// <param name="numPixels"></param>
            /// <returns></returns>
            public bool SetRowHeight(int rowNum, double height)
            {
                // Input is NOT in pixels.  It's a HEIGHT value that seems consistent across all spreadsheet versions.
                try
                {
                    SP.ExcelWorksheet.Rows[rowNum].RowHeight = height;
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Gets the row height in points (72 points/inch)
            /// </summary>
            /// <param name="rowNum"></param>
            /// <returns></returns>
            public double GetRowHeight(int rowNum)
            {
                try
                {
                    return ((double)SP.ExcelWorksheet.Rows.EntireRow[rowNum].RowHeight);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (0);
                }
            }

            /// <summary>
            /// Sets the column width in points (72 points/inch)
            /// </summary>
            /// <param name="colNum"></param>
            /// <param name="width"></param>
            /// <returns></returns>
            public bool SetColumnWidth(int colNum, double width)
            {
                // Input is NOT in pixels.  It's a WIDTH value that seems consistent across all spreadsheet versions.
                try
                {
                    SP.ExcelWorksheet.Columns[colNum].ColumnWidth = width;
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Gets the column width in points (72 points/inch)
            /// </summary>
            /// <param name="colNum"></param>
            /// <returns></returns>
            public double GetColumnWidth(int colNum)
            {
                try
                {
                    return ((double)SP.ExcelWorksheet.Columns.EntireColumn[colNum].ColumnWidth);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (0);
                }
            }



            /// <summary>
            /// Auto-fits range of cells based on the amount of text in each cell.  NOTE: Should be invoked ONLY AFTER all text has been entered.
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            internal bool AutoFitRows(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Rows.AutoFit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Auto-fits range of cells based on the amount of text in each cell.  NOTE: Should be invoked ONLY AFTER all text has been entered.
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            internal bool AutoFitColumns(int startRow, int startCol, int endRow, int endCol)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Columns.AutoFit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Sets the range of cells backcolor
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="endCol"></param>
            /// <param name="color"></param>
            /// <returns></returns>
            public bool Backcolor(int startRow, int startCol, int endRow, int endCol, Color color)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, endCol);
                    range.Interior.Color = ColorTranslator.ToOle(color);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Sets the cell backcolor
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <param name="color"></param>
            /// <returns></returns>
            public bool Backcolor(int row, int col, Color color)
            {
                return (Backcolor(row, col, row, col, color));
            }
            /// <summary>
            /// Places a box around the perimeter of the specified range of cells.
            /// </summary>
            /// <param name="startRow"></param>
            /// <param name="startCol"></param>
            /// <param name="endRow"></param>
            /// <param name="EndCol"></param>
            /// <param name="style"></param>
            /// <param name="weight"></param>
            /// <param name="color"></param>
            /// <returns></returns>
            public bool Box(int startRow, int startCol, int endRow, int EndCol, LineStyle style, BorderWeight weight, Color color)
            {
                try
                {
                    Excel.Range range = SP.SetRange(startRow, startCol, endRow, EndCol);
                    range.BorderAround2((Excel.XlLineStyle)style, (Excel.XlBorderWeight)weight, Excel.XlColorIndex.xlColorIndexNone, ColorTranslator.ToOle(color));
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Places a box around the perimeter of the specified cell.
            /// </summary>
            /// <param name="row"></param>
            /// <param name="col"></param>
            /// <param name="style"></param>
            /// <param name="weight"></param>
            /// <param name="color"></param>
            /// <returns></returns>
            public bool Box(int row, int col, LineStyle style, BorderWeight weight, Color color)
            {
                return (Box(row, col, row, col, style, weight, color));
            }
        }
        #endregion

        #region PageSetupClass
        public class PageSetupClass 
        {
            Spreadsheet SP;

            public PageSetupClass (Spreadsheet sp)
            {
                SP = sp;
            }

            /// <summary>
            /// Puts spreadsheet in landscape mode
            /// </summary>
            /// <returns></returns>
            public bool Landscape()
            {
                return (PageOrientation(Excel.XlPageOrientation.xlLandscape));
            }

            /// <summary>
            /// Puts spreadsheet in portrait mode
            /// </summary>
            /// <returns></returns>
            public bool Portrait()
            {
                return (PageOrientation(Excel.XlPageOrientation.xlPortrait));
            }

            /// <summary>
            /// Selects spreadsheet orientation
            /// </summary>
            /// <param name="orientation"></param>
            /// <returns></returns>
            private bool PageOrientation(Excel.XlPageOrientation orientation)
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Worksheet w = SP.ExcelWorksheet;
                    w.PageSetup.Orientation = orientation;
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Fits spreadsheet to N pages wide
            /// </summary>
            /// <param name="numWide"></param>
            /// <returns></returns>
            public bool FitToPagesWide(int numWide)
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Worksheet w = SP.ExcelWorksheet;
                    w.PageSetup.Zoom = false;
                    w.PageSetup.FitToPagesWide(numWide);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Fits spreadsheet to N pages tall.
            /// </summary>
            /// <param name="numTall"></param>
            /// <returns></returns>
            public bool FitToPagesTall(int numTall)
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Worksheet w = SP.ExcelWorksheet;
                    w.PageSetup.Zoom = false;
                    w.PageSetup.FitToPagesTall(numTall);
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Sets spreadsheet margins in points (72 points/inch)
            /// </summary>
            /// <param name="top"></param>
            /// <param name="bottom"></param>
            /// <param name="left"></param>
            /// <param name="right"></param>
            /// <param name="header"></param>
            /// <param name="footer"></param>
            /// <returns></returns>
            private bool Margins(double top, double bottom, double left, double right, double header, double footer)
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Worksheet w = SP.ExcelWorksheet;
                    w.PageSetup.BottomMargin = bottom;
                    w.PageSetup.TopMargin = top;
                    w.PageSetup.LeftMargin = left;
                    w.PageSetup.RightMargin = right;
                    w.PageSetup.HeaderMargin = header;
                    w.PageSetup.FooterMargin = footer;
                    return (true);
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Sets "standard" narrow spreadsheet margins
            /// </summary>
            /// <returns></returns>
            public bool MarginsNarrow()
            {
                try
                {
                    return (Margins(54, 54, 18, 18, 22, 22));  // 72 picas per inch
                                                               //return (Margins(0.75, 0.75, 0.25, 0.25, 0.3, 0.3));
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }

            /// <summary>
            /// Sets "standard" normal spreadsheet margins
            /// </summary>
            /// <returns></returns>
            public bool MarginsNormal()
            {
                try
                {
                    return (Margins(54, 54, 50, 50, 22, 22));  // 72 picas per inch
                                                               //return (Margins(0.75, 0.75, 0.7, 0.7, 0.3, 0.3));
                }
                catch (Exception ex)
                {
                    SP.LastException = ex.ToString();
                    return (false);
                }
            }
        }
#endregion

#region Fileclass
        public class FileClass
        {
            Spreadsheet SP;

            public FileClass(Spreadsheet sp)
            {
                SP = sp;
            }

            /// <summary>
            /// Saves spreadsheet with specfied filename
            /// </summary>
            /// <param name="saveAsPath"></param>
            /// <param name="DeleteFirst"></param>
            /// <returns></returns>
            public bool SaveAs(string saveAsPath, bool DeleteFirst)
            {
                try
                {
                    if (DeleteFirst)
                    {
                        if (System.IO.File.Exists(saveAsPath))
                        {
                            System.IO.File.Delete(saveAsPath);
                        }
                    }
                    SP.ExcelApp.DisplayAlerts = true;
                    SP.ExcelWorkbook.SaveAs(
                        saveAsPath,
                        Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                        Type.Missing,
                        Type.Missing,
                        false,
                        false,
                        Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                        Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        true);
                    SP.ExcelApp.DisplayAlerts = true;
                    return (true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR: Couldn't complete save operation: " + ex.ToString());
                    return (false);
                }
            }
        }
#endregion

    }
}
