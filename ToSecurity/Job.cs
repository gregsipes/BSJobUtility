using BSJobBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace ToSecurity
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "ToSecurity";
            JobDescription = "Creates an Excel file of employee records with title and address";
            AppConfigSectionName = "ToSecurity";
        }

        public override void ExecuteJob()
        {
            try
            {
                //get records from database
                WriteToJobLog(JobLogMessageType.INFO, "Getting data");
                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.SBSReports, "Proc_Select_ToSecurity").ToList();
               

                if (results != null && results.Count() > 0)
                {
                    WriteToJobLog(JobLogMessageType.INFO, "Creating spreadsheet");


                    //create the Excel Interop reference and workbook
                    Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook workbook;
                    excel.Application.Workbooks.Add();
                    workbook = excel.Application.ActiveWorkbook;
                    excel.Application.DisplayAlerts = false;
                    Microsoft.Office.Interop.Excel.Worksheet activeWorksheet = workbook.Sheets[1];

                    //build the column headers from the first result
                    int headerCounter = 1;
                    foreach (KeyValuePair<string, object> field in results[0])
                    {

                        activeWorksheet.Cells[1, headerCounter] = field.Key;

                        int number;
                        if (int.TryParse(field.Value.ToString(), out number))
                            activeWorksheet.Columns[headerCounter].NumberFormat = "@";  
                        
                        headerCounter++;
                    }

                    int rowCounter = 2;

                    foreach (Dictionary<string, object> result in results)
                    {
                        int fieldCounter = 1;

                        foreach (KeyValuePair<string, object> field in result)
                        {
                            DateTime dateTime;
                            if (DateTime.TryParse(field.Value.ToString(), out dateTime))
                                activeWorksheet.Cells[rowCounter, fieldCounter] = dateTime.ToShortDateString();
                            else
                                activeWorksheet.Cells[rowCounter, fieldCounter] = field.Value.ToString();


                            fieldCounter++;
                        }

                        rowCounter++;
                    }

                    //delete the previous version of the file
                    if (File.Exists(GetConfigurationKeyValue("OutputFileName")))
                    {
                        WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, $"Deleting existing file at {GetConfigurationKeyValue("OutputFileName")}");
                        File.Delete(GetConfigurationKeyValue("OutputFileName"));
                    }

                    activeWorksheet.Columns.AutoFit();

                    //save the new version
                    workbook.SaveAs(Filename: GetConfigurationKeyValue("OutputFileName"));
                    WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, $"New file created at {GetConfigurationKeyValue("OutputFileName")}");
                    workbook.Close(SaveChanges: false);

                    excel.Application.Quit();
                    excel.Quit();
                    ReleaseExcelObject(workbook);
                    ReleaseExcelObject(excel);

                    workbook = null;
                    excel = null;
                }

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in releasing object :" + ex);
                obj = null;
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
