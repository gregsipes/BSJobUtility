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
        private JobBase _managedJob;

        public JobExecutor(string jobName, string[] args)
        {
            _jobName = jobName;

            // setup job
            SetupJob();

            // pre execution
            PreExecution(args);

            // execute job
            ExecuteJob();

            // post execution
            PostExecution();
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
            else if (_jobName == "ManifestLoadAdvance")
                _managedJob = new ManifestLoadAdvance.Job();
            else if (_jobName == "PBSInvoiceExportLoad")
                _managedJob = new PBSInvoiceExportLoad.Job();
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
            _managedJob.PostExecuteJob();
        }

        public void Dispose()
        {
            if (_managedJob != null)
                _managedJob = null;
        }
    }
}
