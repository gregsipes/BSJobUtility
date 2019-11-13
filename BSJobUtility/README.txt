BSJobUtility

Application Name: BS Job Utility

Authors
	Greg Sipes		PCA Technology Group	gsipes@pcatg.com
	Paul Buckley	Buffalo News			pbuckley@buffnews.com

Description
	The intention of this application is to provide a consolidated 
    development platform for automation jobs. The jobs are designed as 
    individual projects within the solution. The jobs are executed by passing 
    a parameter on the command-line.

    Example: BSJobUtility.exe /j ParkingPayroll 

    Additional job specific arguments can be passed in if the job supports 
    additional or optional arguments, see the job specific README.

Deployment Instructions
    Each job within the BS Job Utility is intended to be executed by a 
    scheduler or manually using the /j command-line argument. To execute a job 
    simply execute BSJobUtility.exe passing the /j switch followed by 
    the job name. The BS Job Utility will report back an exit code 
    upon completion. This exit code will allow any calling applications to 
    determine whether the job executed successfully or not. If the utility is 
    being run manually, the exit code will be displayed to the user. An exit 
    code of 0 indicates a successful run.  

	Example Deployment
    Let's say for example there are jobs ParkingPayroll and PBSMacroLoad in the BS Job 
    Utility and each job needs to be executed by a job scheduler. To do this 
    one would place the BS Job Utility someplace where the scheduler 
    can access it then create 2 jobs in the scheduler. Each job would call 
    the BSJobUtility.exe passing the /j switch with the job name. 
    Any configurable settings would be maintained in the common .config file 
    located with the exe file. All logging would be performed by the BS 
    Job Utility and stored in the database. The only logging needed at the 
    scheduler level would be general job execution logging.

Notes
    By default, job log entries will expire after 27 weeks. At this point, when 
    the job log cleanup job is run, expired entries will be deleted. This 
    behavior can be overridden by creating a configuration in an individual 
    job's section called "LogEntryAgeLimitWeeks". The value must be an integer 
    representing the number of weeks a log entry will be allowed to remain.
    Entries that are not a positive non-zero integer will be ignored and the 
    default value will be used.

Development Notes
    When creating entries in the config file, special characters, such as "&", 
    "<", ">", "'", and """ should be entered as "&amp;", "&lt;", "&gt;", 
    "&apos;", and "&quot;" respectively.

    Each job should inherit from BSJobBase.JobBase. Any features that 
    could be used by other jobs or future jobs should be added to JobBase.

    Each job should contain a README file. The template for the job README file
    is in the BSJobBase project.

Change Log
	Build 1.0.0 
		-Initial Release