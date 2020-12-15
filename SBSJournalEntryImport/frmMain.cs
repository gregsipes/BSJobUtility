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
//
// Currently, anyone with SBSReports execution privileges can run this app.

namespace SBSJournalEntryImport
{
    public partial class FrmMain : Form
    {

        #region Declarations
        // Constants
        const string JobName = "SBSJournalEntryImport";

        // Class declarations

        // Other global stuff
        ActiveDirectory UserInfo;
        VersionStatusBar StatusBar;
        string WorkingFolder = "";
        int NumSpreadsheetRows = 0;
        int NumSpreadsheetCols = 0;

        #endregion

        #region Initialization
        public FrmMain()
        {
            InitializeComponent();

            // Job log start
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job starting", JobName);

            // Get configuration values.  

            WorkingFolder = Config.GetConfigurationKeyValue("SBSJournalEntryImport", "WorkingFolder");

            // Create event handlers if any 
            //EXAMPLE: TxtAddressLine1.TextChanged += new System.EventHandler(this.VendorTxtBox_TextChanged);

            // Menu strip initialization (where needed)
            //MainMenuStrip.Renderer = new CustomMenuStripRenderer();

            // Get the current (logged-in) username.  It will be in the form DOMAIN\username.
            //   Terminate if the user is not credentialed to run this app.
            UserInfo = new ActiveDirectory();
            BSOUClass bsouSBS = UserInfo.BSOUList.Find(x => x.Credential.ToLower() == "bsou_sbsreports");
            BSOUClass bsouAdmin = UserInfo.BSOUList.Find(x => x.Credential.ToLower() == "bsadmin");
            if ((bsouSBS is null) && (bsouAdmin is null))
            {
                BroadcastError("You do not have the appropriate credentials (BSOU_SBSReports) to run this app.", null);
                System.Environment.Exit(1);
            }

            // Add status bar (2 segment default, with version)
            StatusBar = new VersionStatusBar(this);

        }
        #endregion

        #region SQL
        public static SqlDataReader SQLQuery(string qryName, SqlParameter[] orderParams)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.SBSJournalEntryImport,
                CommandType.StoredProcedure,
                qryName,
                orderParams);
            return (rdr);
        }

        public static SqlDataReader SQLQuery(string qryName, CommandType command)
        {
            SqlDataReader rdr = DataIO.ExecuteQuery(
                Enums.DatabaseConnectionStringNames.ISInventory,
                command,
                qryName);
            return (rdr);
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

        private int ConfirmSpreadsheetFormat(Spreadsheet import)
        {
            // Confirm that the spreadsheet has the appropriate format.  It MUST be exactly as was used
            //   to define the import table tblJournalEntries_Imported.

            // First, get the fields and datatypes for the import table in column order
            const int MAX_HEADER_ROWS = 20;
            int FirstDataRow = 0;
            try
            {
                string query = "SELECT column_name, data_type FROM information_schema.columns WHERE table_name = 'tblJournalEntries_Imported'";
                List<Dictionary<string, object>> Fields = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.SBSJournalEntryImport, CommandType.Text, query);

                // Item 0 must be the full general ledger ID, all numeric.  It should also be the first time a numeric is found in this column.
                //   walk through the spreadsheet until we find this (or until the maximum number of header rows is hit, in which case we have an error).
                for (int h = 1; h <= MAX_HEADER_ROWS; h++)
                {
                    string v = import.GetCellValue(h, 1).ToString();
                    try
                    {
                        bool int64okay = Int64.TryParse(v, out Int64 FullGL);
                        if (int64okay)
                        {
                            FirstDataRow = h;
                            break;
                        }
                    }
                    catch
                    {
                        // try next row
                    }
                }
                if (FirstDataRow == 0)
                {
                    BroadcastError("ERROR:  Unable to correctly parse spreadsheet:  Header is > " + MAX_HEADER_ROWS.ToString() + " rows OR the first column is not in the correct format (it should be the GL ID value, all numeric)", null);
                    return (-1);
                }

                // Next, check that every column in the header is correct and that there are exactly the correct number of columns

                for (int c = 0; c < Fields.Count; c++)
                {

                    object imprt = import.GetCellValue(FirstDataRow, c + 1);

                    string cell = (imprt is null) ? "" : imprt.ToString();
                    string columnname = Fields[c]["column_name"].ToString();
                    string datatype = Fields[c]["data_type"].ToString().ToLower();
                    switch (datatype)
                    {
                        case "varchar":
                            // Nothing to do here unless it's the Amount column - then confirm that it looks like Money
                            if (columnname.ToLower() == "amount") // should be column 11 (one-based)
                            {
                                if (!IsCurrency(cell))
                                {
                                    BroadcastError("ERROR:  Column " + (c + 1).ToString() + " (Amount) is not a currency value; it contains '" + cell + "'.  Spreadsheet does not have the correct format as previous spreadsheets.  Cannot continue", null);
                                    return (-1);
                                }
                            }
                            break;
                        case "datetime":
                            bool dateokay = DateTime.TryParse(cell, out DateTime d);
                            if (!dateokay)
                            {
                                BroadcastError("ERROR:  Column " + (c + 1).ToString() + " (" + columnname + ") is not a date; it contains '" + cell + "'.  Spreadsheet does not have the correct format as previous spreadsheets.  Cannot continue", null);
                                return (-1);
                            }
                            break;
                        case "int":
                            bool intokay = int.TryParse(cell, out int n);
                            if (!intokay)
                            {
                                BroadcastError("ERROR:  Column " + (c + 1).ToString() + " (" + columnname + ") is not an integer; it contains '" + cell + "'.  Spreadsheet does not have the correct format as previous spreadsheets.  Cannot continue", null);
                                return (-1);
                            }
                            break;
                        default:
                            BroadcastError("ERROR:  Column " + (c + 1).ToString() + " has data type " + datatype + " which is not correct. The import table does not have the correct schema.  Cannot continue", null);
                            return (-1);
                    }
                }

                // There Should be fields.count columns
                NumSpreadsheetRows = import.GetNumRows();
                NumSpreadsheetCols = import.GetNumColumns();
                if (NumSpreadsheetCols > 0)
                {
                    if (NumSpreadsheetCols != Fields.Count)
                    {
                        BroadcastError("ERROR:  Spreadsheet has " + NumSpreadsheetCols.ToString() + " columns.  It is supposed to have " + Fields.Count.ToString() + ".  Cannot continue", null);
                        return (-1);
                    }
                }

                // Success!  Return with the first data row.
                return (FirstDataRow);
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying to confirm spreadsheet format", ex);
                return (-1);
            }
        }

        private bool IsCurrency(string cellval)
        {
            // Currency-formatted strings can only have the following character values:
            //   0-9 $ ( ) - , and space
            // TBD replace with RegEx

            for (int i = 0; i < cellval.Length; i++)
            {
                char c = cellval[i];
                if (!(c >= '0' && c <= '9') || (c == '$') || (c == '(') || (c == ')') || (c == '-') || (c == ',') || (c == ' '))
                {
                    return (false);
                }
            }
            return (true);
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

        #region Events
        private void TxtImportFile_Click(object sender, EventArgs e)
        {
            CmdBrowse_Click(sender, e);
        }

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
                    StatusBar.AddText(0, "Successfully loaded " + TxtImportFile.Text);
                }
                else
                {
                    TxtImportFile.Text = "";
                    CmdImport.Enabled = false;
                    StatusBar.AddText(0, "File not loaded");
                }
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying to browse for file", ex);
                StatusBar.AddText(0, "ERROR");
                return;
            }
        }

        private void CmdExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void CmdImport_Click(object sender, EventArgs e)
        {
            // Convert the selected Excel file to tab-delimited, CSV format.
            //   This will convert any currency values > 999.99 to double-quoted strings, with comma separators in the thousands (and millions) 

            int numrows = 0;
            try
            {
                // Open the spreadsheet
                StatusBar.AddText(0, "Converting spreadsheet to tab-delimited file...");
                Spreadsheet import = new Spreadsheet(TxtImportFile.Text);

                // Confirm that it conforms to table design.  The return value will either be the starting data row OR -1 if failure.
                int FirstDataRow = ConfirmSpreadsheetFormat(import);
                if (FirstDataRow == -1)
                {
                    StatusBar.AddText(0, "SPREADSHEET PARSING ERROR");
                    return;
                }
                else
                {
                    DataIO.WriteToJobLog(Enums.JobLogMessageType.INFO, "Spreadsheet " + TxtImportFile.Text + " passes consistency checks", JobName);
                }

                // Save as tab-delimited file (delete any identically-named file first)
                string CSVFilename = WorkingFolder + " SBSJournalEntry_" + UserInfo.Username + ".txt";
                import.File.SaveAs(CSVFilename, Spreadsheet.FileFormat.TabDelimited, true);
                import.Terminate();

                // Truncate, then bulk insert the CSV file into table tblJournalEntries_Imported
                //   IMPORTANT:  The stored procedure assumes that the first N rows of the file are header rows.  This MUST be specified here.

                StatusBar.AddText(0, "Loading file...");
                int NumHeaderRowsInFile = FirstDataRow - 1;
                SqlParameter[] BulkParams = new SqlParameter[2];
                BulkParams[0] = new SqlParameter("strCSVFilename", CSVFilename);
                BulkParams[1] = new SqlParameter("strNumHeaderRows", NumHeaderRowsInFile.ToString());
                SqlDataReader rdr = SQLQuery("Proc_BulkInsertSBSJournalEntries", BulkParams);

                // If the returned Sum(Amount) is nonzero, generate a message to warn the user before proceeding
                // NOTE Do a read here to collect any Exception error from the query:
                //   rdr.Read();
                // Because there's a possibility of a bulk read error there wasn't a known way to consolidate
                //   the next queries into the above procedure, so we put them here to collect the data we need
                //   regarding consistency checks on the successfully-imported data.  As a once-a-month process there isn't much need for optimization.

                string query = "SELECT Sum(PARSE([Amount] AS MONEY)) AS SumAmount FROM  tblJournalEntries_Imported";
                List<Dictionary<string, object>> Fields = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.SBSJournalEntryImport, CommandType.Text, query);
                bool AmountOkay = double.TryParse(Fields[0]["SumAmount"].ToString(), out double TotalAmount);
                if (TotalAmount != 0)
                {
                    DialogResult result = MessageBox.Show("Sum(Amount) for this spreadsheet is " + TotalAmount.ToString("C") + ". Okay to continue?", "NON-ZERO TOTAL SUM", MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes)
                    {
                        DataIO.WriteToJobLog(Enums.JobLogMessageType.WARNING, "Job halted by user for nonzero total amount on spreadsheet (" + TotalAmount.ToString("C") + ")", JobName);
                        StatusBar.AddText(0, "ERROR");
                        return;
                    }
                }

                // Confirm that we read all file lines (or 1 less than the file if an extraneous LF was inserted during Accounting's load into the spreadsheet).
                query = "SELECT count(*) AS Count FROM  tblJournalEntries_Imported";
                Fields = DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.SBSJournalEntryImport, CommandType.Text, query);

                numrows = (int)Fields[0]["Count"];
                if (NumSpreadsheetRows - numrows > 1)
                {
                    // Generate this error only if we differ by more than one line.
                    BroadcastError("ERROR trying import spreadsheet " + TxtImportFile.Text, null);
                    StatusBar.AddText(0, "ERROR");
                    return;
                }

            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying import spreadsheet " + TxtImportFile.Text, ex);
                StatusBar.AddText(0, "ERROR");
                return;
            }

            // Transfer the data from the imported table to the target table (tblJournalEntries),
            //   performing any data translations along the way.  
            //   As of 12/10/2020 the only known translation is Amount (text) to currency.
            // We also have to overwrite any existing identical records lest we create duplicates

            // "identical as defined by the accounting department is identical values in the following fields
            //     Bal Type
            //     Period
            //     Year
            //     Month
            //     Quarter
            // These are the same for any import.  If these records are in the target table and an import contains the same values then
            //     we have to delete these records from the target table first so that only the new import data is retained.

            try
            {
                StatusBar.AddText(0, "Copying data to Journal Entry database...");

                DataIO.ExecuteNonQuery(Enums.DatabaseConnectionStringNames.SBSJournalEntryImport, CommandType.StoredProcedure, "Proc_Insert_SBSJournalEntries");
            }
            catch (Exception ex)
            {
                BroadcastError("ERROR trying to copy imported data to database", ex);
                StatusBar.AddText(0, "ERROR");
            }

            BroadcastInfo("Successfully imported spreadsheet " + TxtImportFile.Text + ": " + numrows.ToString() + " records imported.", null);
            StatusBar.AddText(0, "Done!");
        }

        private void FrmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job completed", JobName);
        }
        #endregion


    }
}
