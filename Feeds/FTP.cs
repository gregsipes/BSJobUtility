using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Feeds
{
   public class FTP
    {
        private string Host { get; set; }

        private string UserName { get; set; }

        private string Password { get; set; }

        public FTP(string host, string userName, string password)
        {
            Host = host;
            UserName = userName;
            Password = password;
        }

        public void UploadFile(FileInfo sourceFile, string destinationPath)
        {
            using (WebClient webClient = new WebClient())
            {
                webClient.Credentials = new NetworkCredential(UserName, Password);
                webClient.UploadFile($"ftp://{Host}/{destinationPath}/{sourceFile.Name}", WebRequestMethods.Ftp.UploadFile, sourceFile.FullName);
            }
        }

        public void CreateDirectory(string path)
        {
            WebRequest request = WebRequest.Create($"ftp://{Host}/{path}");
            request.Method = WebRequestMethods.Ftp.MakeDirectory;
            WebResponse response = request.GetResponse();
        }

        public bool CheckIfDirectoryExists(string path)
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create($"ftp://{Host}/{path}");
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                return true;
            }
            catch (WebException ex)
            {
                return false;
            }
        }
    }
}
