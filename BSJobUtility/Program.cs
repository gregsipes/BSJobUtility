using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSJobUtility
{
    class Program
    {
        static int Main(string[] args)
        {
            string jobName = "";

            // determine from command line arguments which job to execute
            if (args.Count() == 0)
            {
                Console.WriteLine("No commandline arguments supplied, nothing to do.");
                PrintCommandLineHelp();
                Console.WriteLine("Exit code:" + "0");
                return 0;
            }

            try
            {
                for (int i = 0; i < args.Length; i++)
                {
                    //help 
                    if (args[i] == "/h")
                    {
                        PrintCommandLineHelp();
                        Console.WriteLine(string.Format("Exit code: {0}", 0));
                        return 0;
                    } //run job
                    else if (args[i] == "/j")
                        jobName = args[i + 1];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error encountered parsing commandline arguments. " + ex.Message);
                int exitCode = ex.HResult;
                Console.WriteLine(string.Format("Exit code: {0}", exitCode.ToString()));
                return exitCode;
            }

            if (jobName == "")
                Console.WriteLine("Unable to determine job name, nothing to do.");
            else
            {
                int exitCode = ExecuteJob(jobName, args);
                Console.WriteLine(string.Format("Exit code: {0}", exitCode.ToString()));
                return exitCode;
            }


            Console.WriteLine(string.Format("Exit code: {0}", 0));

            Console.ReadLine();

            return 0;

        }

        private static void PrintCommandLineHelp()
        {
            string tab = new string(' ', 4);

            Console.WriteLine();
            Console.WriteLine("Usage: ");
            Console.WriteLine(tab + "/h" + tab + "Show this help information.");
            Console.WriteLine(tab + "/j" + tab + "Number of job to execute.");
            Console.WriteLine();
            Console.WriteLine("Example: BSJobUtility.exe /j ParkingPayroll");
            Console.WriteLine();
        }

        private static int ExecuteJob(string jobName, string[] args)
        {
            try
            {
                JobExecutor jobExecutor = new JobExecutor(jobName, args);

                if (jobExecutor != null)
                    jobExecutor.Dispose();

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error encountered while executing job. {0} ", ex.Message);
                return ex.HResult;
            }

        }
    }
}

