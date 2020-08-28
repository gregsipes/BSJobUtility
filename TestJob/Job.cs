﻿using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace TestJob
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            try
            {

               WriteToJobLog(JobLogMessageType.INFO, "Test job is running");

                throw new Exception("Testing...");

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }

        public override void SetupJob()
        {
            JobName = "Test Job";
            JobDescription = "This job is meant for testing only. ";
            AppConfigSectionName = "TestJob";
        }

    }
}
