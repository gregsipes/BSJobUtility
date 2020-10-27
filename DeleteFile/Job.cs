using BSJobBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static BSGlobals.Enums;

namespace DeleteFile
{
    public class Job : JobBase
    {

        public override void SetupJob()
        {
            JobName = "Delete File";
            JobDescription = "Deletes files from directories based on age.";
            AppConfigSectionName = "DeleteFile";
        }

        public override void ExecuteJob()
        {
            try
            {
                //get the files to search for
                List<Dictionary<string, object>> results = ExecuteSQL(DatabaseConnectionStringNames.BSJobUtility, "Proc_Select_Delete_Files").ToList();

                foreach (Dictionary<string, object> result in results)
                {
                    List<string> files = Directory.GetFiles(result["Path"].ToString(), result["FileSearchPattern"].ToString()).ToList();

                    //delete each file
                    foreach(string file in files)
                    {
                        //todo: check if is last file

                        File.Delete(file);
                    }

                    //todo: check subdirectories
                }

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }


    }
}
