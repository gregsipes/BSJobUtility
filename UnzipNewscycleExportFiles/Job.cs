﻿using BSJobBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;

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

                // Get the name (w/o extension) of this zip file's root folder and delete this folder if it exists and delete any old folders.
                bool UnzipOkay = true;
                string FolderName = Path.GetFileNameWithoutExtension(zf);
                try
                {
                    using (ZipArchive archive = ZipFile.OpenRead(zf))
                    {
                        List<ZipArchiveEntry> ListOfZipFolders = archive.Entries.Where(x => x.FullName.EndsWith("/")).ToList();
                        // There should be at least one folder in this list of folders.  Pick off the root folder name and use it to create the target directory
                        string ZipFolderPathname = ListOfZipFolders[0].FullName;
                        string[] ZipFolderRootname = ZipFolderPathname.Split('/');
                        FolderName = ZipFolderRootname[0];
                    }
                }
                catch (Exception ex)
                {
                    SendMail($"Error in Job: {JobName}", "Unable to open/get Newscycle EXPORT root path from zip file " + zf + ": " + ex.ToString(), false);
                    WriteToJobLog(JobLogMessageType.ERROR, "Unable to open/get Newscycle EXPORT root path from zip file " + zf + ": " + ex.ToString());
                    // Try it with the default folder name, so keep going and see what happens...
                }

                DirectoryExists = Directory.Exists(SourceFolder + FolderName);
                if (DirectoryExists)
                {
                    try
                    {
                        Directory.Delete(SourceFolder + FolderName, true);
                    }
                    catch (Exception ex)
                    {
                        SendMail($"Error in Job: {JobName}", "Unable to delete Newscycle EXPORT data folder " + SourceFolder + FolderName + " " + ex.ToString(), false);
                        WriteToJobLog(JobLogMessageType.ERROR, "Unable to delete Newscycle EXPORT data folder " + SourceFolder + FolderName + " " + ex.ToString());
                        UnzipOkay = false;
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
                        SendMail($"Error in Job: {JobName}", "Unable to unzip Newscycle EXPORT data folder " + zf + " " + ex.ToString(), false);
                        WriteToJobLog(JobLogMessageType.ERROR, "Unable to unzip Newscycle EXPORT data folder " + zf + " " + ex.ToString());
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
                            SendMail($"Error in Job: {JobName}", "Unable to delete Newscycle EXPORT Touch folder " + TouchFolder + " " + ex.ToString(), false);
                            WriteToJobLog(JobLogMessageType.ERROR, "Unable to delete Newscycle EXPORT Touch folder " + TouchFolder + " " + ex.ToString());
                            UnzipOkay = false;
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
                        SendMail($"Error in Job: {JobName}", "Unable to delete Newscycle EXPORT zip file " + zf + " " + ex.ToString(), false);
                        WriteToJobLog(JobLogMessageType.ERROR, "Unable to delete Newxcycle EXPORT zip file " + zf + " " + ex.ToString());
                        UnzipOkay = false;
                    }
                }

                if (UnzipOkay)
                {
                    WriteToJobLog(JobLogMessageType.INFO, "Unzip of Newscycle EXPORT file " + zf + " successfully completed");
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
