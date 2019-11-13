ParkingPayroll

Job Information
    Name: Parking Payroll

    Description: Updates employees and departments from SBS.

    Configuration Section: ParkingPayroll.

	Configuration Keys:
	    RemoteServerInstance        SBS Payroll server database instance for use in remote server query.
        RemoteDatabaseName          SBS Payroll server database name.
        RemoteUserName				SBS Payroll server databaser user name.
		RemotePassword				SBS Payroll server database password.

    Required Additional Arguments: None.

    Optional Additional Arguments: None.

Change Log

     01/31/17  v1.0.7  DAH
	   -   Changed modGeneral.
	
	 08/06/14  v1.0.6  DAH
	   -   Added modStartupVerification.
	   -   Changed modGeneral, modGlobals & modParkingPayroll.
	
	 09/18/13  v1.0.5  DAH
	   -   Changed modParkingPayroll.
	
	 10/24/12  v1.0.4  DAH
	   -   Changed modGeneral & modParkingPayroll.
	
	 08/28/12  v1.0.3  Req #2649  DAH
	   -   Set AutomaticAppStatus in gobjAppStatus.
	
	 08/28/12  v1.0.3  Req #2646  DAH
	   -   Added reference to BSEmail3.dll.
	   -   Deleted reference to BSEmail.dll.
	   -   Changed coding related to new email reference.
	
	 08/02/12  v1.0.2  Req #2631  DAH
	   -   Changed modParkingPayroll.
	
	 07/25/12  v1.0.1  DAH
	   -   Changed modGeneral & modParkingPayroll.
	
	 03/25/08  v1.0.0  Req #751  DAH
	 04/05/06  v1.0.0  DAH
	   -   Application created to execute stored procedure specified in [Program Options]
	       in .ini file, updating database with information from SBS.
	   -   Application based on ParkingLawson.
