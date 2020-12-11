using BSJobBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using static BSGlobals.Enums;

namespace UnzipNewscycleExportFiles
{
    public class Job : JobBase
    {
        public override void ExecuteJob()
        {
            // Confirm that we have access to the export folder (default:  \\circfs\backup)

            string SourceFolder = GetConfigurationKeyValue("sourcefolder");
            bool DirectoryExists = Directory.Exists(SourceFolder);
            if (!DirectoryExists)
            {
                // Directory could not be accessed (or does not exist).  Log an error and exit

                SendMail($"Error in Job: {JobName}", "Unable to access Newscycle EXPORT folder " + SourceFolder, false);
                WriteToJobLog(JobLogMessageType.ERROR, "Unable to access Newscycle EXPORT folder " + SourceFolder);
                Environment.Exit(1);
            }

            // Check for any zip file(s). There should typically be only one but we can loop here
            //   to process any and all zip files.

            string ZipFileExtension = GetConfigurationKeyValue("compressedfileextension");
            string[] ZipFiles = Directory.GetFiles(SourceFolder, "*." + ZipFileExtension);
            List<string> ZipFileList = ZipFiles.ToList();
            foreach (string zf in ZipFileList)
            {
                WriteToJobLog(JobLogMessageType.INFO, "Unzip of Newscycle EXPORT file " + zf + " started");

                // First things first.  Make a copy of the zip file; overwrite any older one.  This will allow us to manually
                //   retry any failed extraction (and subsequent data update) should we suffer some kind of failure.

                try
                {
                    File.Copy(zf, zf + "_COPY", true);
                } 
                catch (Exception ex)
                {
                    WriteToJobLog(JobLogMessageType.WARNING, "Unable to make a copy of Newscycle EXPORT zip file " + zf + ": " + ex.ToString());
                }

                // Get the name (w/o extension) of this zip file's root folder and delete this folder if it exists and delete any old folders.
                bool UnzipOkay = true;
                bool WarningsGiven = false;

                string FolderName = Path.GetFileNameWithoutExtension(zf);
                try
                {
                    using (ZipArchive archive = ZipFile.OpenRead(zf))
                    {
                        List<ZipArchiveEntry> ListOfZipFolders = archive.Entries.Where(x => x.FullName.EndsWith("/")).ToList();
                        // There *should* be at least one folder in this list of folders.  Pick off the root folder name and use it to create the target directory
                        if (ListOfZipFolders.Count > 0)
                        {
                            string ZipFolderPathname = ListOfZipFolders[0].FullName;
                            string[] ZipFolderRootname = ZipFolderPathname.Split('/');
                            if (ZipFolderRootname.Length > 0)
                            {
                                FolderName = ZipFolderRootname[0];
                            }
                            else
                            {
                                FolderName = ZipFolderPathname;
                            }
                        }
                        else
                        {
                            WriteToJobLog(JobLogMessageType.INFO, "Unable to find Newscycle EXPORT root path from zip file " + zf + " (no folders found, using " + FolderName + " as the default root path)");
                        }
                    }
                }
                catch (Exception ex)
                {
                    SendMail($"Warning from Job: {JobName}", "Unable to open/get Newscycle EXPORT root path from zip file " + zf + ": " + ex.ToString(), false);
                    WriteToJobLog(JobLogMessageType.WARNING, "Unable to open/get Newscycle EXPORT root path from zip file " + zf + ": " + ex.ToString());
                    WarningsGiven = true;
                    // Try it with the default folder name, so keep going and see what happens...
                }

                bool DeleteErrorOccurred = false;
                DirectoryExists = Directory.Exists(SourceFolder + FolderName);
                if (DirectoryExists)
                {
                    try
                    {
                        Directory.Delete(SourceFolder + FolderName, true);
                    }
                    catch (Exception ex)
                    {
                        // Interestingly enough, even with the Delete method's 2nd argument set to true (recursively delete all folders)
                        //   we can still generate the exception "directory not empty".  When this happens, delay, then try one more time before generating 
                        //   an error message.
                        // Also, don't fail simply because we couldn't perform a delete.  Let the zip extraction go ahead anyway and let it 
                        //   generate a fatal exception if it can't extract.
                        WriteToJobLog(JobLogMessageType.WARNING, "Unable to delete Newscycle EXPORT data folder (attempt #1)" + SourceFolder + FolderName + " " + ex.ToString());
                        WarningsGiven = true;
                        try
                        {
                            System.Threading.Thread.Sleep(2000); // Wait 2 seconds before trying this again.
                            Directory.Delete(SourceFolder + FolderName, true);
                        }
                        catch (Exception ex1)
                        {
                            DeleteErrorOccurred = true;
                            WriteToJobLog(JobLogMessageType.ERROR, "Unable to delete Newscycle EXPORT data folder (attempt #2)" + SourceFolder + FolderName + " " + ex1.ToString());
                            //UnzipOkay = false;
                        }
                    }
                }

                // Unzip the zip file. This will decompress all data files as well as the DumpControl files.
                if (UnzipOkay)
                {
                    try
                    {
                       ZipFile.ExtractToDirectory(zf, SourceFolder);
                    }
                    catch (Exception ex)
                    {
                        string DeleteFailureMsg = "";
                        if (DeleteErrorOccurred)
                        {
                            DeleteFailureMsg = " (EXPORT data folder could not be deleted prior to extraction): ";
                        }
                        SendMail($"Error in Job: {JobName}", "Unable to unzip Newscycle EXPORT data folder " + zf + DeleteFailureMsg + ex.ToString(), false);
                        WriteToJobLog(JobLogMessageType.ERROR, "Unable to unzip Newscycle EXPORT data folder " + zf + DeleteFailureMsg + ex.ToString());
                        UnzipOkay = false;
                    }
                }

                // Within the SourceFolder (or one of its subfolders) should ALSO be the Touch file that must be extracted last.
                //   Extracting this file will set off the import chain of apps running as SQL jobs every 5 minutes.
                //   There should only be a single touch file (as long as cleanup is working okay).

                if (UnzipOkay)
                {
                    string TouchFolder = SourceFolder + FolderName + "\\Touch";
                    DirectoryExists = Directory.Exists(TouchFolder);
                    if (DirectoryExists)
                    {
                        try
                        {
                            Directory.Delete(TouchFolder, true);
                        }
                        catch (Exception ex)
                        {
                            SendMail($"Warning in Job: {JobName}", "Unable to delete Newscycle EXPORT Touch folder " + TouchFolder + " " + ex.ToString(), false);
                            WriteToJobLog(JobLogMessageType.WARNING, "Unable to delete Newscycle EXPORT Touch folder " + TouchFolder + " " + ex.ToString());
                            WarningsGiven = true;
                            //UnzipOkay = false;
                        }
                    }
                }

                if (UnzipOkay)
                {
                    string[] TouchFiles = Directory.GetFiles(SourceFolder + FolderName, "*." + ZipFileExtension, SearchOption.AllDirectories);
                    try
                    {
                        string TargetFolder = SourceFolder + FolderName + "\\";
                        ZipFile.ExtractToDirectory(TouchFiles[0], TargetFolder);
                    }
                    catch (Exception ex)
                    {
                        SendMail($"Error in Job: {JobName}", "Unable to unzip Newscycle EXPORT Touch folder " + TouchFiles[0] + " " + ex.ToString(), false);
                        WriteToJobLog(JobLogMessageType.ERROR, "Unable to unzip Newxcycle EXPORT Touch folder " + TouchFiles[0] + " " + ex.ToString());
                        UnzipOkay = false;
                    }
                }

                // All files have been successfully extracted.  Delete the main zip file
                if (UnzipOkay)
                {
                    try
                    {
                        File.Delete(zf);
                    }
                    catch (Exception ex)
                    {
                        SendMail($"WARNING in Job: {JobName}", "Unable to delete Newscycle EXPORT zip file " + zf + " " + ex.ToString(), false);
                        WriteToJobLog(JobLogMessageType.WARNING, "Unable to delete Newxcycle EXPORT zip file " + zf + " " + ex.ToString());
                        WarningsGiven = true;
                        //UnzipOkay = false;
                    }
                }

                if (UnzipOkay)
                {
                    string Warnings = "";
                    if (WarningsGiven)
                    {
                        Warnings = " (with Warnings)";
                    }
                    WriteToJobLog(JobLogMessageType.INFO, "Unzip of Newscycle EXPORT file " + zf + " successfully completed" + Warnings);
                }
                else
                {
                    WriteToJobLog(JobLogMessageType.INFO, "Unzip of Newscycle EXPORT file " + zf + " unsuccessful");
                }

            } // foreach (string zf in ZipFileList)

            // TBD After all unzip processing has completed, any additional cleanup goes here.

        }

        public override void SetupJob()
        {
            JobName = "Unzip Newscycle Export Files";
            JobDescription = @"Extract Newscycle files";
            AppConfigSectionName = "UnzipNewscycleExportFiles";
        }
    }
}
