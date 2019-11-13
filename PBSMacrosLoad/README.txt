PBSMacrosLoad

Job Information
    Name: PBS Macros Load

    Description: Moves files from \\circ\spoolcmro to \\Synergy\SERops\To Be Loaded\PBS ASCII. A file
				only get moved if it contains an '=' sign, and either doesn't exist in the loads table or
				exists but has a different last modified timestamp

    Configuration Section: PBSMacrosLoad.

	Configuration Keys:


    Required Additional Arguments: None.

    Optional Additional Arguments: None.

Change Log

     01/31/17  v1.2.3  DAH
	   -   Changed modGeneral.
	
	 03/18/15  v1.2.2  DAH
	   -   Changed modPBS2MacroLoad.
	
	 08/06/14  v1.2.1  DAH
	   -   Added modStartupVerification.
	   -   Changed modGeneral, modGlobals & modPBS2MacroLoad.
	
	 11/21/13  v1.2.0  DAH
	   -   Changed modPBS2MacroLoad.
	
	 09/18/13  v1.1.9  DAH
	   -   Changed modPBS2MacroLoad.
	
	 10/24/12  v1.1.8  DAH
	   -   Changed modGeneral & modPBS2MacroLoad.
	
	 09/17/12  v1.1.7  Req #2646  DAH
	   -   Added reference to BSEmail3.dll.
	   -   Deleted reference to BSEmail.dll.
	   -   Changed coding related to new email reference.
	
	 08/02/12  v1.1.6  Req #2631  DAH
	   -   Changed modPBS2MacroLoad.
	
	 07/24/12  v1.1.5  DAH
	   -   Changed modGeneral & modPBS2MacroLoad.
	
	 03/30/10  v1.1.4  Req #1733  DAH
	   -   Changed modGeneral & modPBS2MacroLoad.
	
	 03/11/10  v1.1.3  Req #1701  DAH
	   -   Changed modGeneral, modGlobals & modPBS2MacroLoad.
	
	 02/09/10  v1.1.2  Req #1596  DAH
	   -   Added reference to BSLoad.dll.
	   -   Changed modGeneral, modGlobals & modPBS2MacroLoad.
	   -   Deleted frmInProgress.
	
	 03/26/09  v1.1.1  Req #1099  DAH
	   -   Changed modGeneral, modGlobals & modPBS2MacroLoad.
	
	 11/20/07  v1.1.0  DAH
	   -   Added reference to BSEmail.dll.
	   -   Removed reference to ASPEmail.dll.
	   -   Changed modGeneral, modGlobals & modPBS2MacroLoad.
	
	 07/19/05  v1.0.4  DAH
	   -   Changed modGeneral.
	
	 12/15/04  v1.0.3  DAH
	   -   Changed modGeneral.
	
	 03/11/04  v1.0.2  DAH
	   -   Changed modGeneral & modPBS2MacroLoad.
	
	 12/29/03  v1.0.1  DAH
	   -   Changed modPBS2MacroLoad.
	
	 10/06/03  v1.0.0
	   -   Created app to copy files created in PBS for loading
	       into Macrofiche.
