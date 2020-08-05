using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using AppStatusControl;
using BSGlobals;

namespace AppStatusMonitor
{
    public partial class frmMain : Form
    {

        int NumActivitiesPerMonitor = 15;
        int LookbackInDays = 90;
        const string JobName = "App Status Monitor";
        List<string> AppNameList = new List<string>();
        List<AppStatusUserControl> StatusMonitorList = new List<AppStatusUserControl>();
        int TimerUpdateIntervalInMsec = 10000; // A default value for the timer update interval, which is read from config on startup
        DateTime StartupTime = DateTime.Now;
        Size MonitorSize = new Size(0, 0);
        bool SingleLineMode = true;

        public frmMain()
        {
            InitializeComponent();

            // Get the refresh interval in seconds and convert to msec.
            bool activitycountokay = int.TryParse(Config.GetConfigurationKeyValue("AppStatusMonitor", "NumActivitiesPerMonitor"), out NumActivitiesPerMonitor);
            bool lookbackokay = int.TryParse(Config.GetConfigurationKeyValue("AppStatusMonitor", "LookbackInDays"), out LookbackInDays);
            bool success = int.TryParse(Config.GetConfigurationKeyValue("AppStatusMonitor", "UpdateIntervalInSecs"), out int result);
            if (success)
            {
                TimerUpdateIntervalInMsec = 1000 * result;
            }
            timUpdateStatus.Enabled = true;

            SetPanelSize();
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, "Job starting", JobName);
        }

        private void SetPanelSize()
        {
            pnlMonitors.Size = new Size(this.ClientSize.Width - 8, this.ClientSize.Height - 8);
        }

        private void timUpdateStatus_Tick(object sender, EventArgs e)
        {
            try
            {
                // First time in:  Initial timer value is 1 msec so that we get a fast initial data load.
                //   Afterward, set the timer update interval to whatever was read from the config file
                if (timUpdateStatus.Interval != TimerUpdateIntervalInMsec)
                {
                    GetAppNames(true);  // First time in, set this to true so we load app names all the way back to LookbackInSeconds
                    timUpdateStatus.Interval = TimerUpdateIntervalInMsec;
                }
                else
                {
                    // Get the names of all apps in the log table (for the last Tick seconds)
                    GetAppNames(false);  // Al subsequent calls:  Set this to false so we load app names only since the last timer tick.
                }

                // For each app, get the last N cycles and check for errors originating from within the app.
                int NumCycles = NumActivitiesPerMonitor;

                SqlParameter[] ActivityParams = new SqlParameter[3];
                for (int i = 0; i < AppNameList.Count; i++)
                {
                    //  command.Parameters.Add(new SqlParameter("@MessageType", type.ToString("d")));
                    ActivityParams[0] = new SqlParameter("@pvchrJobName", AppNameList[i]);
                    ActivityParams[1] = new SqlParameter("@pvintLookbackInDays", LookbackInDays);
                    ActivityParams[2] = new SqlParameter("@pvintNumCycles", NumCycles);
                    SqlDataReader rdr = DataIO.ExecuteQuery(Enums.DatabaseConnectionStringNames.EventLogs, CommandType.StoredProcedure, "Proc_Select_Last_N_Activities", ActivityParams); // A misnamed sproc.  Should be N, not 5
                    UpdateMonitor(rdr, StatusMonitorList[i]);
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.ERROR, "Tick error: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error in timer tick: " + ex.ToString());
            }
        }


        private void GetAppNames(bool firstTime)
        {
            try
            {
                bool MonitorsNeedRecreating = false;

                // Get a list of all apps in the event log that have run in the past N days
                // Syntax:
                //   Results is a list of dictionary entries of type <string>,<object> as required by ExecuteSQL.
                //   For each dictionary entry, 
                //        <string> will contain the field name
                //        <object> will contrain the value for that field (which must be explictly typed later)
                //   Each entry in the list represents a single row from the stored procedure.

                string sprocname = "";
                SqlParameter param;
                if (firstTime)
                {
                    // First time in:  Collect app names all the way back to the Lookback interval
                    sprocname = "dbo.Proc_Select_List_Of_All_Apps";
                    param = new SqlParameter("@pvintLookbackInDays", Config.GetConfigurationKeyValue("AppStatusMonitor", "LookbackInDays"));
                }
                else
                {
                    // All subsequent queries:  Collect app names since the last timer tick
                    sprocname = "dbo.Proc_Select_List_Of_All_New_Apps";
                    param = new SqlParameter("@pvintLookbackInSeconds", Config.GetConfigurationKeyValue("AppStatusMonitor", "UpdateIntervalInSecs"));
                }
                List<Dictionary<string, object>> results =
                    DataIO.ExecuteSQL(Enums.DatabaseConnectionStringNames.EventLogs,
                    sprocname,
                    param);

                foreach (Dictionary<string, object> entry in results)
                {                    // <object> will be the AppName, once it's converted to a string.
                    string appname = ((string)entry["JobName"]);

                    // Check if this name is already on the app list.  If not,
                    //    Add it to the list
                    //    Mark that monitors need to be recreated.
                    if (!AppNameList.Contains(appname))
                    {
                        AppNameList.Add(appname);
                        MonitorsNeedRecreating = true;
                    }
                }

                // If any monitor needs to be created, then
                //    Delete all existing monitors
                //    Recreate the monitor list in sort order

                if (MonitorsNeedRecreating)
                {
                    DeleteAllMonitors();
                    AppNameList.Sort();
                    foreach (string name in AppNameList)
                    {
                        CreateMonitor(name);
                    }
                    ArrangeMonitors();
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "GetAppNames: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to get app names: " + ex.ToString());
            }
        }

        private void ArrangeMonitors()
        {
            try
            {
                // Arrange the monitors to fit within the frame (with scrolling if necessary)
                // What's the width of the panel interior and the monitor control?
                //   And how many controls can we fit within the panel's width?
                int panelx = pnlMonitors.Width;
                int numpanelsacross = panelx / MonitorSize.Width;
                if (numpanelsacross == 0)
                {
                    numpanelsacross = 1;
                }

                // Separate the panels vertically as well
                int numpanelsdown = (StatusMonitorList.Count + (numpanelsacross - 1)) / numpanelsacross;
                for (int i = 0; i < numpanelsdown; i++)
                {
                    for (int j = 0; j < numpanelsacross; j++)
                    {
                        if (i * numpanelsacross + j < StatusMonitorList.Count)
                        {
                            AppStatusUserControl uc = StatusMonitorList[j + i * numpanelsacross];
                            uc.Left = j * MonitorSize.Width;
                            uc.Top = i * MonitorSize.Height;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "Arrange Monitors: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to arrange monitors: " + ex.ToString());
            }
        }

        private void CreateMonitor(string name)
        {
            try
            {
                // Create an application monitor, and render it visible
                AppStatusUserControl uc = new AppStatusUserControl(NumActivitiesPerMonitor, SingleLineMode)
                {
                    AppName = name,
                    Visible = true
                };
                StatusMonitorList.Add(uc);
                this.pnlMonitors.Controls.Add(uc);

                // Monitor size is the same for all monitors. Save it for later use.
                MonitorSize = new Size(uc.Width, uc.Height);

                // Add a mouseclick event handler so we can use it to toggle between display modes
                //uc.ucMouse_Click += new EventHandler((sender, e) => ucMouse_Click(sender, e));
                uc.ucMouse_Click += ucMouse_Click;

            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "Create Monitor: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to create monitor: " + ex.ToString());
            }
        }

        private void DeleteAllMonitors()
        {
            // Destroy all monitors
            try
            {
            //foreach (AppStatusUserControl uc in pnlMonitors.Controls)  Can't use this approach because we're deleting controls and will skip some as the control list compresses
            for (int i = pnlMonitors.Controls.Count - 1; i >= 0; i--)
            {
                if (pnlMonitors.Controls[i] is AppStatusUserControl)
                {
                    AppStatusUserControl uc = (AppStatusUserControl)pnlMonitors.Controls[i];
                    pnlMonitors.Controls.Remove(uc);
                    //uc.Dispose(); // Is this needed?  
                }
            }
            StatusMonitorList.Clear();
            pnlMonitors.Refresh();

            } catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "DeleteAllMonitors: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to delete all monitors: " + ex.ToString());
            }
        }

        private void UpdateMonitor(SqlDataReader rdr, AppStatusUserControl appStatusUserControl)
        {
            // Update the selected data monitor

            // The SQL data reader passed into this routine should have 3 datasets attached to it:
            //   - A list of the last N dates (or fewer) of this app's acitivity that was other than "started/completed"
            //   - A list of all warnings and errors from the jobs that ran during any of those dates
            //   - The app's very last starting/completed message to determine if the app is still running

            try
            {
                // First result:  The list of the last N activity dates
                List<DateTime> ActivityDates = new List<DateTime>();
                while (rdr.Read())
                {
                    ActivityDates.Add((DateTime)rdr["LogDate"]);
                }

                // Second result:  The list of all warnings and errors (containing LogDate, MessageType and Message)
                List<IssuesType> IssuesList = new List<IssuesType>();
                rdr.NextResult();
                while (rdr.Read())
                {
                    IssuesType issue = new IssuesType
                    {
                        LogDate = (DateTime)rdr["LogDate"],
                        MessageType = (int)rdr["MessageType"],
                        Message = rdr["Message"].ToString()
                    };
                    IssuesList.Add(issue);
                }

                // Third result:  The app's last starting or completed message.  This will be either zero or one record in length
                bool AppIsRunning = false;
                DateTime LastExecutionTime = new DateTime(1900, 01, 01);
                rdr.NextResult();
                while (rdr.Read())
                {
                    AppIsRunning = (rdr["Message"].ToString() == "Job starting") ? true : false;
                    LastExecutionTime = (DateTime)rdr["LogDate"];
                }
                rdr.Close();

                Color color = (AppIsRunning) ? Color.White : Color.Blue;
                appStatusUserControl.SetLEDColor(AppStatusUserControl.LEDs.LEDActivity, 0, color);

                // Determine which activities had errors or warnings
                List<LEDStatusesType> LEDStatuses = ComputeLEDStatuses(appStatusUserControl, ActivityDates, IssuesList);

                // and light the appropriate leds the appropriate color.
                for (int i = 0; i < LEDStatuses.Count; i++)
                {
                    appStatusUserControl.SetLEDColor(AppStatusUserControl.LEDs.LEDStatus, i, LEDStatuses[i].LEDColor);
                }
                appStatusUserControl.ClearLEDs(LEDStatuses.Count); // This clears (turns off) any remaining LEDs.

                // Set the current runtime value to the last execution time in the log

                appStatusUserControl.RunTime = LastExecutionTime;
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "Failed to correctly update monitor " + appStatusUserControl.AppName + ": " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to update monitor: " + ex.ToString());
            }

        }

        private List<LEDStatusesType> ComputeLEDStatuses(AppStatusUserControl appStatusUserControl, List<DateTime> activityDates, List<IssuesType> issuesList)
        {
            // Take the list of issues and bounce them across the list of activity dates to determine in which date range the issue arose.
            //   Return the appropriate LED color for each activity date
            // TBD - This can be optimized.  Old LED statuses will never change; only newer ones and any that are currently ongoing.  
            //   We should take advantage of that to eliminate the redundant efforts to rebuild LED statuses with every update.

            List<LEDStatusesType> LEDStatusList = new List<LEDStatusesType>();
            for (int i = 0; i < activityDates.Count; i++)
            {
                LEDStatusList.Add(new LEDStatusesType());
            }

            try
            {
                // Find out which activity this issue belongs to
                for (int j = 0; j < issuesList.Count; j++)
                {
                    IssuesType issue = issuesList[j];
                    for (int i = 0; i < activityDates.Count - 1; i++)
                    {
                        // is it activity [i]?
                        string Messages = "";
                        if ((issue.LogDate <= activityDates[i]) && (issue.LogDate >= activityDates[i + 1]))
                        {
                            // Why yes it is!  Set the activity's LED to either yellow (if it was green) or red (unconditionally) based on the message type.  
                            Messages = activityDates[i].ToShortDateString() + " " + activityDates[i].ToShortTimeString();
                            issue.LEDNum = i;
                            issuesList[j] = issue;  // Save this; we'll use it when hovering over a LED
                            LEDStatusesType ledstatus = LEDStatusList[i];
                            switch (issue.MessageType)
                            {
                                case 1:
                                    ledstatus.LEDColor = Color.Green;
                                    break;
                                case 2:
                                    if (ledstatus.LEDColor == Color.Green)
                                    {
                                        ledstatus.LEDColor = Color.Yellow;
                                    }
                                    break;
                                case 3:
                                    ledstatus.LEDColor = Color.Red;
                                    break;
                                default:
                                    break;
                            }

                            // Save the message as well...
                            Messages += "\r\n" + issue.Message;  // TBD .Messages is Unnecessary, get rid of it in the class
                            // And save the status message back to the control for later tool tipping
                            LEDStatusList[i] = ledstatus;  // TBD Unnecessary
                            appStatusUserControl.SetLEDMessage(i, Messages);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ComputeLEDStatuses: " + ex.ToString(), appStatusUserControl.AppName);
                // TBD throw new Exception("Error trying to update LED status: " + ex.ToString());
            }
            return (LEDStatusList);
        }

        private class IssuesType
        {
            public DateTime LogDate { get; set; }
            public int MessageType { get; set; }
            public string Message { get; set; }
            public int LEDNum { get; set; }

            public IssuesType()
            {
                LogDate = DateTime.Now;
                MessageType = 0;
                Message = "";
                LEDNum = -1;
            }
        }

        private class LEDStatusesType
        {
            public Color LEDColor { get; set; }
            public LEDStatusesType()
            {
                LEDColor = Color.Green;
            }
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DateTime StopTime = DateTime.Now;
            double ElapsedTime = ((TimeSpan)(StopTime - StartupTime)).TotalSeconds;
            TimeSpan t = TimeSpan.FromSeconds(ElapsedTime);
            string result = string.Format("{0:D2}h:{1:D2}m:{2:D2}.{3:D3}s", t.Hours, t.Minutes, t.Seconds, t.Milliseconds);
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, "Runtime: " + result, JobName);
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, "Job completed", JobName);
        }

        private void frmMain_ResizeEnd(object sender, EventArgs e)
        {
            // At the end of a form resize, redistribute the existing monitors
            SetPanelSize();
            ArrangeMonitors();
        }

        private void AppStatusMonitor_Hover(object sender, EventArgs e)
        {
            // TBD THIS IS OBSOLETE (NEVER HIT AND NOT NEEDED)
            // Mouse just hovered over a LED.  Get the LED's index and the name of the app that triggered this event
            AppStatusUserControl uc = (AppStatusUserControl)sender;
            int lednum = uc.LEDNum;
            string appname = uc.AppName;
            string msg = uc.GetLEDMessage(lednum);
        }

        private void ucMouse_Click(object sender, EventArgs e)
        {
            ToggleDisplay();
        }

        private void pnlMonitors_Click(object sender, EventArgs e)
        {
            ToggleDisplay();
        }

        private void ToggleDisplay()
        {
            // Toggle the display between single line and multiline
#if false
            // TBD This will be the preferred way - faster
            for (int i = 0; i < StatusMonitorList.Count; i++)
            {
                AppStatusUserControl uc = StatusMonitorList[i];
                uc.ToggleDisplayMode();
            }
#else
            // This needs to be optimized as above - this method has way more overhead
            SingleLineMode = !SingleLineMode;
            DeleteAllMonitors();
            AppNameList.Sort();
            foreach (string name in AppNameList)
            {
                CreateMonitor(name);
            }
            ArrangeMonitors();
#endif
        }
    }
}
