BSJobBase

Job Information

    Description: 	The intention of this project is to provide a base class to centralize any code that can be shared amongst the derived classes. The intention is
					that each of the derived classes will be a separate SQL job. The base class provided must be inherited and cannot be directly implemented.

					7/20/2020 - A few months into building this solution, we decided to add the first UI project. As a result, to avoid changing each job,
					we decided to move the majority of the code in this class to the BSGlobals project to share with the UI's. In some cases, it adds an extra
					hop, but in other cases will are useful. Additionally, as each new job project is added to the solution, it will need to reference both 
					the BSGlobals project as well as this BSJobBase project. The only reason that a project needs to reference BSGlobals is for the enums, 
					ALL other calls should be made through the base or derived classes.

Change Log

    Build 1.0.0
        - Initial release.





	