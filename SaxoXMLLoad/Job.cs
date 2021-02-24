using BSJobBase;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using static BSGlobals.Enums;

namespace SaxoXMLLoad
{
    public class Job : JobBase
    {
        public override void SetupJob()
        {
            JobName = "Saxo XML Load";
            JobDescription = "TODO";
            AppConfigSectionName = "SaxoXMLLoad";
        }

        public override void ExecuteJob()
        {
            try
            {
                List<string> files = Directory.GetFiles(GetConfigurationKeyValue("InputDirectory"), "*.xml").ToList();

                if (files != null && files.Count() > 0)
                {
                    foreach (string file in files)
                    {
                        FileInfo fileInfo = new FileInfo(file);

                        if (fileInfo.Length > 0) //ignore empty files
                        {

                            Dictionary<string, object> previouslyLoadedFile = ExecuteSQL(DatabaseConnectionStringNames.Newshole, "dbo.Proc_Select_SaxoXML_Loads_If_Processed",
                                                                new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind))).FirstOrDefault();


                            if (previouslyLoadedFile == null)
                            {
                                //make sure we the file is no longer being edited
                                if ((DateTime.Now - fileInfo.LastWriteTime).TotalMinutes > Int32.Parse(GetConfigurationKeyValue("SleepTimeout")))
                                {
                                    WriteToJobLog(JobLogMessageType.INFO, $"{fileInfo.FullName} found");
                                    CopyAndProcessFile(fileInfo);
                                }
                            }
                            //else
                            //{
                            //    ExecuteNonQuery(DatabaseConnectionStringNames.Newshole, "Proc_Insert_Loads_Not_Loaded",
                            //                    new SqlParameter("@pvchrOriginalDir", fileInfo.Directory.ToString()),
                            //                    new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                            //                    new SqlParameter("@pdatLastModified", fileInfo.LastWriteTime),
                            //                    new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                            //                    new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                            //                    new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString()));
                            //}
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                LogException(ex);
                throw;
            }
        }


        private void CopyAndProcessFile(FileInfo fileInfo)
        {
            string backupFileName = GetConfigurationKeyValue("OutputDirectory") + fileInfo.Name + "_" + DateTime.Now.ToString("yyyyMMddhhmmsstt") + ".xml";
            Int32 loadsId = 0;


            //copy file to backup directory
            File.Copy(fileInfo.FullName, backupFileName, true);
            WriteToJobLog(JobLogMessageType.INFO, "File copied to " + backupFileName);

            //update or create a load id
            Dictionary<string, object> result = ExecuteSQL(DatabaseConnectionStringNames.Newshole, "Proc_Insert_SaxoXML_Loads",
                                                                                        new SqlParameter("@pvchrOriginalDirectory", fileInfo.Directory.ToString() + "\\"),
                                                                                        new SqlParameter("@pvchrOriginalFile", fileInfo.Name),
                                                                                        new SqlParameter("@pdatLastModified", new DateTime(fileInfo.LastWriteTime.Year, fileInfo.LastWriteTime.Month, fileInfo.LastWriteTime.Day, fileInfo.LastWriteTime.Hour, fileInfo.LastWriteTime.Minute, fileInfo.LastWriteTime.Second, fileInfo.LastWriteTime.Kind)),
                                                                                        new SqlParameter("@pvchrNetworkUserName", System.Security.Principal.WindowsIdentity.GetCurrent().Name),
                                                                                        new SqlParameter("@pvchrComputerName", System.Environment.MachineName.ToLower()),
                                                                                        new SqlParameter("@pvchrLoadVersion", Assembly.GetExecutingAssembly().GetName().Version.ToString())).FirstOrDefault();
            loadsId = Int32.Parse(result["saxoxml_loads_id"].ToString());
            WriteToJobLog(JobLogMessageType.INFO, $"Loads ID: {loadsId}");


            XmlReader reader = XmlReader.Create(backupFileName);

            reader.ReadToFollowing("head");

            if (reader.Name == "head")
            {
                XElement headerNode = (XElement)XElement.ReadFrom(reader);
                XElement pageNode = headerNode.Element("pageplanningsystem");
                XElement issueNode = headerNode.Element("issue");

                ExecuteNonQuery(DatabaseConnectionStringNames.Newshole, "Proc_Update_SaxoXML_Loads_Info",
                                        new SqlParameter("@pintSaxoXMLLoadsID", loadsId),
                                        new SqlParameter("@pvchrPubYYYYMMDD", issueNode.Element("date").Attribute("value").Value.ToString()),
                                        new SqlParameter("@pvchrPublication", pageNode.Element("name").Attribute("value").Value.ToString()),
                                        new SqlParameter("@pvchrXMLVersion", headerNode.Element("xmlversion").Attribute("value").Value.ToString()));


                //reader.ReadToFollowing("pagedescription");
            }

            while (!reader.EOF)
            {

                if (reader.Name != "pagedescription")
                    reader.ReadToFollowing("pagedescription");

                if (!reader.EOF)
                {
                    XElement node = (XElement)XElement.ReadFrom(reader);

                    //add the page record to the database
                    ExecuteNonQuery(DatabaseConnectionStringNames.Newshole, "Proc_Insert_SaxoXML_Pages",
                                            new SqlParameter("@pintSaxoXMLLoadsID", loadsId),
                                            new SqlParameter("@pintPageNumber", node.Element("pagenumber").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pvchrSection", node.Element("section").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pintSectionPageNumber", node.Element("sectionpagenumber").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pvchrCategory", node.Element("category").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pvchrChannel", node.Element("channel").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pvchrTemplateFile", node.Element("templatefile").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pvchrZone", node.Element("zone").Attribute("value").Value.ToString()));


                    //add ad records to the database
                    XElement adsGroupNode = node.Element("ads"); //parent container

                    List<XElement> adNodes = adsGroupNode.Elements("ad").ToList();

                    foreach (XElement adNode in adNodes)
                    {
                        ExecuteNonQuery(DatabaseConnectionStringNames.Newshole, "Proc_Insert_SaxoXML_Ads",
                                            new SqlParameter("@pintSaxoXMLLoadsID", loadsId),
                                            new SqlParameter("@pintPageNumber", node.Element("pagenumber").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pintAdID", adNode.Element("file").Attribute("name").Value.ToString().Replace(".pdf", "")),
                                            new SqlParameter("@pvchrDescription", adNode.Element("description").Attribute("value").Value.ToString()),
                                            new SqlParameter("@pvchrFileName", adNode.Element("file").Attribute("name").Value.ToString()),
                                            new SqlParameter("@pnumX1InMM", adNode.Element("pos").Attribute("x1").Value.ToString()),
                                            new SqlParameter("@pnumX2InMM", adNode.Element("pos").Attribute("x2").Value.ToString()),
                                            new SqlParameter("@pnumY1InMM", adNode.Element("pos").Attribute("y1").Value.ToString()),
                                            new SqlParameter("@pnumY2InMM", adNode.Element("pos").Attribute("y2").Value.ToString()));
                    }



                }

            }


            WriteToJobLog(JobLogMessageType.INFO, "About to execute Proc_Insert_Update_Brainworks_Accounts_SaxoXML");
            ExecuteNonQuery(DatabaseConnectionStringNames.Newshole, "Proc_Insert_Update_Brainworks_Accounts_SaxoXML",
                                        new SqlParameter("@pintSaxoXMLLoadsID", loadsId),
                                        new SqlParameter("@pvchrBrainworksServiceInstance", GetConfigurationKeyValue("RemoteServerName")),
                                        new SqlParameter("@pvchrBrainworksDatabase", GetConfigurationKeyValue("RemoteDatabaseName")),
                                        new SqlParameter("@pvchrUserName", GetConfigurationKeyValue("RemoteUserName")),
                                        new SqlParameter("@pvchrPassword", GetConfigurationKeyValue("RemotePassword")));


            WriteToJobLog(JobLogMessageType.INFO, "About to execute Proc_Update_SaxoXML_Loads_Successful");
            ExecuteNonQuery(DatabaseConnectionStringNames.Newshole, "Proc_Update_SaxoXML_Loads_Successful", new SqlParameter("@pintSaxoXMLLoadsID", loadsId));
            WriteToJobLog(JobLogMessageType.INFO, "Load information updated");

        }
    }
}
