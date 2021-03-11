using BSJobBase;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSJobUtility
{
    public class JobExecutor : IDisposable
    {
        private readonly string _jobName;
        private readonly string _group;
        private readonly string _version;
        private JobBase _managedJob;
        private bool _wasSuccessful;

        public JobExecutor(string jobName, string group, string version, string[] args)
        {
            try
            {
                _jobName = jobName;
                _group = group;
                _version = version;

                // setup job
                SetupJob();

                // pre execution
                PreExecution(args);

                // execute job
                ExecuteJob();

                _wasSuccessful = true;

            }
            catch(Exception ex)
            {
                _wasSuccessful = false;

                throw new Exception ("JobExecutor exception:  " + ex.ToString()); // 10/01/20 PEB Added Ex to eliminate compiler warning.
            }
            finally
            {
                // post execution
                PostExecution();
            }
        }

        private void SetupJob()
        {
            if (_jobName == "ParkingPayroll")
                _managedJob = new ParkingPayroll.Job();
            else if (_jobName == "PBSMacrosLoad")
                _managedJob = new PBSMacrosLoad.Job();
            else if (_jobName == "CommissionsCreate")
                _managedJob = new CommissionsCreate.Job();
            else if (_jobName == "WrappersLoad")
                _managedJob = new WrappersLoad.Job();
            else if (_jobName == "ManifestLoad")
                _managedJob = new ManifestLoad.Job();
            else if (_jobName == "ManifestFreeLoad")
                _managedJob = new ManifestFreeLoad.Job();
            else if (_jobName == "ManifestLoadAdvance")
                _managedJob = new ManifestLoadAdvance.Job();
            else if (_jobName == "PBSInvoiceExportLoad")
                _managedJob = new PBSInvoiceExportLoad.Job();
            else if (_jobName == "QualificationReportLoad")
                _managedJob = new QualificationReportLoad.Job();
            else if (_jobName == "OfficePay")
                _managedJob = new OfficePay.Job();
            else if (_jobName == "AutoRenew")
                _managedJob = new AutoRenew.Job();
            else if (_jobName == "PressRoomLoad")
                _managedJob = new PressRoomLoad.Job();
            else if (_jobName == "PressRoomFreeLoad")
                _managedJob = new PressRoomFreeLoad.Job();
            else if (_jobName == "PBSInvoiceTotals")
                _managedJob = new PBSInvoiceTotals.Job();
            else if (_jobName == "PBSInvoices")
                _managedJob = new PBSInvoiceLoad.Job();
            else if (_jobName == "UnzipNewscycleExportFiles")
                _managedJob = new UnzipNewscycleExportFiles.Job();
            else if (_jobName == "DMMail")
                _managedJob = new DMMail.Job();
            else if (_jobName == "PayByScanLoadWegmans")
                _managedJob = new PayByScanLoadWegmans.Job();
            else if (_jobName == "PayByScanLoad711")
                _managedJob = new PayByScanLoad711.Job();
            else if (_jobName == "PrepackInsertLoad")
                _managedJob = new PrepackInsertLoad.Job();
            else if (_jobName == "PBSDumpWorkload")
                _managedJob = new PBSDumpWorkload.Job() { GroupName = _group };
            else if (_jobName == "PBSDumpPopulate")
                _managedJob = new PBSDumpPopulate.Job() { GroupName = _group, GroupNumber = _version };
            else if (_jobName == "PBSDumpPost")
                _managedJob = new PBSDumpPost.Job() { GroupName = _group, GroupNumber = _version };
            else if (_jobName == "CircDumpWorkload")
                _managedJob = new CircDumpWorkLoad.Job() { GroupNumber = Convert.ToInt32(_group) };
            else if (_jobName == "CircDumpPopulate")
                _managedJob = new CircDumpPopulate.Job() { GroupNumber = Convert.ToInt32(_group) };
            else if (_jobName == "CircDumpPost")
                _managedJob = new CircDumpPost.Job() { GroupNumber = Convert.ToInt32(_group) };
            else if (_jobName == "SuppliesWorkload")
                _managedJob = new SuppliesWorkload.Job();
            else if (_jobName == "TradeWorkload")
                _managedJob = new TradeWorkload.Job();
            else if (_jobName == "SubBalanceLoad")
                _managedJob = new SubBalanceLoad.Job();
            else if (_jobName == "DeleteFile")
                _managedJob = new DeleteFile.Job();
            else if (_jobName == "DeleteEmptyTMPFiles")
                _managedJob = new DeleteEmptyTMPFiles.Job();
            else if (_jobName == "Feeds")
                _managedJob = new Feeds.Job() { Version = _version };
            else if (_jobName == "PurgeFeeds")
                _managedJob = new PurgeFeeds.Job();
            else if (_jobName == "PackageAssignmentLoad")
                _managedJob = new PackageAssignmentLoad.Job();
            else if (_jobName == "AutoPrintPDF")
                _managedJob = new AutoPrintPDF.Job();
            else if (_jobName == "ExecuteSQL")
                _managedJob = new ExecuteSQL.Job() { Version = _version };
            else if (_jobName == "SBSReportsLoad")
                _managedJob = new SBSReportsLoad.Job();
            else if (_jobName == "SaxoXMLLoad")
                _managedJob = new SaxoXMLLoad.Job();
            else if (_jobName == "TradeAutoEmail")
                _managedJob = new TradeAutoEmail.Job();
            else if (_jobName == "TestJob")
                _managedJob = new TestJob.Job();
            else if (_jobName == "ToSecurity")
                _managedJob = new ToSecurity.Job();
            //else if (_jobName == "RefreshUsers")
            //    _managedJob = new RefreshUsers.Job() { Version = _version };
            else
                throw new Exception("Job name " + _jobName + " is invalid.");

            _managedJob.SetupJob();
        }

        private void PreExecution(string[] args)
        {
            _managedJob.PreExecuteJob(args);
        }

        private void ExecuteJob()
        {
            _managedJob.ExecuteJob();
        }

        private void PostExecution()
        {
            _managedJob.PostExecuteJob(_wasSuccessful);
        }

        public void Dispose()
        {
            if (_managedJob != null)
                _managedJob = null;
        }
    }
}
