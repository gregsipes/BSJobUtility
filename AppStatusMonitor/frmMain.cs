using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AppStatusControl;
using BSGlobals;

// 08-05-20 PEB - V1.0.0.1

namespace AppStatusMonitor
{
    public partial class frmMain : Form
    {

        int NumActivitiesPerMonitor = 15;
        int LookbackInDays = 90;
        const string JobName = "App Status Monitor";
        List<StatusControlType> StatusMonitorList = new List<StatusControlType>();
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
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job starting", JobName);
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
                bool FirstTimeIn = false;
                if (timUpdateStatus.Interval != TimerUpdateIntervalInMsec)
                {
                    FirstTimeIn = true;
                    timUpdateStatus.Interval = TimerUpdateIntervalInMsec;
                }

                // Get the names of all apps in the log table (for the last Tick seconds)
                GetAppNames(StatusMonitorList, FirstTimeIn);  // First time in, this is set to true so we load app names all the way back to LookbackInSeconds

                // For each app, get the last N cycles and check for warnings/errors originating from within the app.
                //   NOTE that N = 1 EXCEPT for the very first time when we will collect enough data for all configured monitor activities.
                //     After the first time, we only need to update the latest activity since older ones will never have updates.
                int NumCycles = (FirstTimeIn) ? NumActivitiesPerMonitor : 1;

                SqlParameter[] ActivityParams = new SqlParameter[2];
                ActivityParams[0] = new SqlParameter("@pvintLookbackInDays", LookbackInDays);
                ActivityParams[1] = new SqlParameter("@pvintMaxActivities", NumCycles);
                SqlDataReader rdr = DataIO.ExecuteQuery(
                    Enums.DatabaseConnectionStringNames.EventLogs,
                    CommandType.StoredProcedure,
                    "Proc_Select_Last_N_Activities_All_Jobs", ActivityParams);

                UpdateMonitors(rdr, StatusMonitorList, FirstTimeIn);

            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.ERROR, "Tick error: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error in timer tick: " + ex.ToString());
            }
        }


        private void GetAppNames(List<StatusControlType> StatusControlList, bool firstTime)
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
                    if (!StatusControlList.Any(n => n.AppName == appname))
                    {
                        StatusControlList.Add(new StatusControlType()
                        {
                            AppName = appname,
                            StatusMonitor = new AppStatusUserControl(NumActivitiesPerMonitor, SingleLineMode)
                        });
                        MonitorsNeedRecreating = true;
                    }
                }

                // If any monitor needs to be created, then
                //    Delete all existing monitors
                //    Recreate the monitor list in sort order

                if (MonitorsNeedRecreating)
                {
                    DeleteAllMonitors(StatusControlList);
                    StatusControlList.OrderBy(x => x.AppName);
                    CreateMonitors(StatusControlList);
                    ArrangeMonitors(StatusControlList);
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "GetAppNames: " + ex.ToString(), JobName);
            }
        }

        private void ArrangeMonitors(List<StatusControlType> StatusControlList)
        {

            try
            {
                // Arrange the monitors to fit within the frame (with scrolling if necessary)
                // What's the width of the panel interior and the monitor control?
                //   And how many controls can we fit within the panel's width?
                // Monitor size is the same for all monitors. Save it for later use.
                MonitorSize = new Size(StatusControlList[0].StatusMonitor.Width, StatusControlList[0].StatusMonitor.Height);
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
                            AppStatusUserControl uc = StatusControlList[j + i * numpanelsacross].StatusMonitor;
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

        private void CreateMonitors(List<StatusControlType> StatusControlList)
        {
            try
            {
                // Create the application monitors

                for (int i = 0; i < StatusControlList.Count; i++)
                {
                    StatusControlType sc = StatusControlList[i];
                    this.pnlMonitors.Controls.Add(sc.StatusMonitor);

                    // Add a mouseclick event handler so we can use it to toggle between display modes
                    //uc.ucMouse_Click += new EventHandler((sender, e) => ucMouse_Click(sender, e));
                    sc.StatusMonitor.ucMouse_Click += ucMouse_Click;
                }

            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "Create Monitor: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to create monitor: " + ex.ToString());
            }
        }

        private void DeleteAllMonitors(List<StatusControlType> StatusControlList)
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
                    }
                }
                pnlMonitors.Refresh();

            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "DeleteAllMonitors: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to delete all monitors: " + ex.ToString());
            }
        }

        private void UpdateMonitors(SqlDataReader rdr, List<StatusControlType> StatusControlList, bool firstTimeIn)
        {
            // Update the selected data monitor

            // The SQL data reader passed into this routine should have 3 datasets attached to it:
            //   - A list of the last N dates (or fewer) of this app's acitivity that was other than "started/completed"
            //   - A list of all warnings and errors from the jobs that ran during any of those dates
            //   - The app's very last starting/completed message to determine if the app is still running

            // TBD Once we are past initialization we need to determine if the query is pointing to the same job start
            //   as the last query.  If not we need to push all activities down one and insert the new one at the beginning
            //   of this job's activity list.

            try
            {
                // First result:  The list of the last N activity dates
                List<LastNRunsType> ActivityDates = new List<LastNRunsType>();
                while (rdr.Read())
                {
                    DateTime js = new DateTime(1900, 01, 01);
                    try
                    {
                        js = (DateTime)rdr["JobStart"];  // A NULL Date will create an exception
                    }
                    catch { }
                    ActivityDates.Add(new LastNRunsType
                    {
                        JobStart = js,
                        JobName = rdr["JobName"].ToString()
                    });
                }

                // Second result:  The list of all warnings and errors (containing LogDate, MessageType and Message)
                List<WarningsErrorsType> IssuesList = new List<WarningsErrorsType>();
                rdr.NextResult();
                while (rdr.Read())
                {
                    DateTime js = new DateTime(1900, 01, 01);
                    try
                    {
                        js = (DateTime)rdr["LogDate"];  // A NULL Date will create an exception
                    }
                    catch { }
                    IssuesList.Add(new WarningsErrorsType
                    {
                        LogDate = js,
                        MessageType = (Enums.JobLogMessageType)rdr["MessageType"],
                        Message = rdr["Message"].ToString(),
                        JobName = rdr["JobName"].ToString()
                    });
                }
                IssuesList = IssuesList.OrderBy(x => x.JobName).ThenBy(y => y.LogDate).ToList();

                // Third result:  The app's last starting or completed message.  This will be either zero or one record in length
                List<LatestMessageType> LatestMessageList = new List<LatestMessageType>();
                rdr.NextResult();
                while (rdr.Read())
                {
                    DateTime js = new DateTime(1900, 01, 01);
                    try
                    {
                        js = (DateTime)rdr["LogDate"];  // A NULL Date will create an exception
                    }
                    catch { }
                    LatestMessageList.Add(new LatestMessageType
                    {
                        Message = rdr["Message"].ToString(),
                        LogDate = js,
                        JobName = rdr["JobName"].ToString()
                    });
                }
                LatestMessageList = LatestMessageList.OrderBy(x => x.JobName).ThenBy(y => y.LogDate).ToList();
                rdr.Close();

                // At this point the JobNames listed in the above lists should be in sync with the StatusControlList passed into this function.
                //   That is, they should all be in sort order.  However, some of the lists read from the database could have missing job names
                //   so we have to carefully go through these lists.
                for (int i = 0; i < StatusControlList.Count; i++)
                {
                    // Determine which activities had errors or warnings
                    StatusControlType uc = StatusControlList[i];
                    List<WarningsErrorsType> AppIssues = new List<WarningsErrorsType>(IssuesList.FindAll(J => J.JobName == uc.AppName));
                    List<LastNRunsType> AppActivities = new List<LastNRunsType>(ActivityDates.FindAll(J => J.JobName == uc.AppName));
                    List<LatestMessageType> AppRunTimeMessage = new List<LatestMessageType>(LatestMessageList.FindAll(J => J.JobName == uc.AppName));

                    // Set LED colors based on no errors (green), warnings (yellow) or errors (red)
                    List<LEDStatusesType> LEDStatuses = ComputeLEDStatuses(uc.StatusMonitor, AppActivities, AppIssues, firstTimeIn);

                    // Set activity LED color and light the appropriate leds the appropriate color.
                    Color color = (LatestMessageList[i].Message.ToLower() == "job starting") ? Color.White : Color.Blue;
                    uc.StatusMonitor.SetLEDColor(AppStatusUserControl.LEDs.LEDActivity, 0, color);

                    // Set the current runtime value to the last execution time in the log
                    uc.StatusMonitor.RunTime = (AppRunTimeMessage.Count > 0) ? AppRunTimeMessage[0].LogDate : new DateTime(1900, 01, 01);
                    uc.StatusMonitor.AppName = uc.AppName;
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "Failed to correctly update monitor: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to update monitor: " + ex.ToString());
            }

        }

        private List<LEDStatusesType> ComputeLEDStatuses(AppStatusUserControl appStatusUserControl, List<LastNRunsType> activityDates, List<WarningsErrorsType> issuesList, bool firstTime)
        {
            // Take the list of issues and bounce them across the list of activity dates to determine in which date range the issue arose.
            //   Return the appropriate LED color for each activity date
            // This has been optimized.  Old LED statuses will never change; only newer ones and any that are currently ongoing.  
            //   We took advantage of that to eliminate the redundant efforts to rebuild LED statuses with every update.
            //   Essentially, the called provides the list activityDates with only the latest activity.

            List<LEDStatusesType> LEDStatusList = new List<LEDStatusesType>();
            try
            {

                // Initialize:  Assume there are no issues associated with each activity date in the list (i.e, color will be black or green)
                //   There are NumActivitiesPerMonitor LEDs to color Black, Green, Yellow or Red
                List<Color> ledcolors = new List<Color>();
                List<string> messages = new List<string>();

                for (int i = 0; i < activityDates.Count; i++)
                {
                    ledcolors.Add(Color.Green);
                    messages.Add(activityDates[i].JobStart.ToShortDateString() + " " + activityDates[i].JobStart.ToShortTimeString());
                }
                for (int i = activityDates.Count; i < NumActivitiesPerMonitor; i++)
                {
                    ledcolors.Add(Color.Black);
                    messages.Add("");
                }

                // The ActivityDates list is in sorted order by activitydate.  Walk through the list
                //  and determine the message and color that should be applied to each specific LED in the list.  

                int lastindex = -1;
                for (int i = 0; i < issuesList.Count; i++)
                {
                    // Find the activity (i.e., the LED index) associated with this issue.  It's the activity whose JobStart value
                    //   is just less than (or equal to) the issue's timestamp.

                    int index = activityDates.FindIndex(x => x.JobStart <= issuesList[i].LogDate);

                    // If this issue is not part of the previous activity, then set the previous activity's message
                    if ((index != lastindex) && (lastindex != -1))
                    {
                        appStatusUserControl.SetLEDMessage(lastindex, messages[lastindex]);
                    }

                    // For any issue, set the LED color accordingly and update the warning/error message for this LED
                    //  NOTE that there can be multiple warning/error messages associated with this LED
                    if (index >= 0)
                    {
                        switch (issuesList[i].MessageType)
                        {
                            case Enums.JobLogMessageType.STARTSTOP:
                                break;  // We don't care about the start/stop messages
                            case Enums.JobLogMessageType.INFO:
                                break;  // We don't care about Info
                            case Enums.JobLogMessageType.WARNING:
                                ledcolors[index] = (ledcolors[index] == Color.Green) ? Color.Yellow : ledcolors[index]; // Turn the LED yellow if still green
                                break;
                            case Enums.JobLogMessageType.ERROR:
                                ledcolors[index] = Color.Red;  // Always turn the LED red on error
                                break;
                            default:
                                break;
                        }
                        messages[index] += "\r\n" + issuesList[i].Message;
                        lastindex = index;

                        // last one
                        if (i == issuesList.Count - 1)
                        {
                            appStatusUserControl.SetLEDMessage(lastindex, messages[lastindex]);
                        }
                    }
                }

                // Optimization:  After the initial load, only the latest activity will ever have updated messages so only update that activity.
                int activitycount = (firstTime) ? NumActivitiesPerMonitor : 1;
                {
                    for (int i = 0; i < activitycount; i++)
                    {
                        appStatusUserControl.SetLEDColor(AppStatusUserControl.LEDs.LEDStatus, i, ledcolors[i]);
                    }
                }
            }
            catch (Exception ex)
            {
                DataIO.WriteToJobLog(Enums.JobLogMessageType.ERROR, "ComputeLEDStatuses: " + ex.ToString(), JobName);
                // TBD throw new Exception("Error trying to update LED status: " + ex.ToString());
            }

            return (LEDStatusList);
        }

        private void ToggleDisplay(List<StatusControlType> StatusControlList)
        {
            // Toggle the display between single line and multiline
            // This will be the preferred way - faster than deleting/recreating the monitors
            for (int i = 0; i < StatusMonitorList.Count; i++)
            {
                AppStatusUserControl uc = StatusMonitorList[i].StatusMonitor;
                uc.ToggleDisplayMode();
            }
            ArrangeMonitors(StatusControlList);
        }

        #region events
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            DateTime StopTime = DateTime.Now;
            double ElapsedTime = ((TimeSpan)(StopTime - StartupTime)).TotalSeconds;
            TimeSpan t = TimeSpan.FromSeconds(ElapsedTime);
            string result = string.Format("{0:D2}h:{1:D2}m:{2:D2}.{3:D3}s", t.Hours, t.Minutes, t.Seconds, t.Milliseconds);
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.INFO, "Runtime: " + result, JobName);
            DataIO.WriteToJobLog(BSGlobals.Enums.JobLogMessageType.STARTSTOP, "Job completed", JobName);
        }

        private void frmMain_ResizeEnd(object sender, EventArgs e)
        {
            // At the end of a form resize, redistribute the existing monitors
            SetPanelSize();
            ArrangeMonitors(StatusMonitorList);
        }

        private void ucMouse_Click(object sender, EventArgs e)
        {
            ToggleDisplay(StatusMonitorList);
        }

        private void pnlMonitors_Click(object sender, EventArgs e)
        {
            ToggleDisplay(StatusMonitorList);
        }

        #endregion

        #region classes

        private class StatusControlType
        {
            public string AppName { get; set; }
            public AppStatusUserControl StatusMonitor { get; set; }

            public StatusControlType()
            {
                AppName = "";
                StatusMonitor = null; // new AppStatusUserControl();
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

        private class LastNRunsType
        {
            public DateTime JobStart { get; set; }
            public string JobName { get; set; }
        }

        private class WarningsErrorsType
        {
            public DateTime LogDate { get; set; }
            public Enums.JobLogMessageType MessageType { get; set; }
            public string Message { get; set; }
            public string JobName { get; set; }
            public int LEDNum { get; set; }
        }

        private class LatestMessageType
        {
            public DateTime LogDate { get; set; }
            public string Message { get; set; }
            public string JobName { get; set; }
        }

        #endregion

    }
}
