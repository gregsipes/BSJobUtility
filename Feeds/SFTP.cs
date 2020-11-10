using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinSCP;

namespace Feeds
{
   public class SFTP
    {
        public Session Session { get; set; }

        public void OpenSession(string host, string user, string password, string fingerprint, string keyFilePath, string keyPassPhrase)
        {
            SessionOptions sessionOptions = new SessionOptions()
                    {
                        Protocol = Protocol.Sftp,
                        HostName = host,
                        UserName = user,
                        Password = password,
                        SshHostKeyFingerprint = fingerprint
                    };

            Session = new Session();

            if (!String.IsNullOrEmpty(keyFilePath))
            {
                sessionOptions.SshPrivateKeyPath = keyFilePath;
                sessionOptions.PrivateKeyPassphrase = keyPassPhrase;
            }

            Session.Open(sessionOptions);
        }

        public bool UploadFile(string sourceFilePath, string destinationFilePath, bool allowResumeSupport, bool allowPreserveTimeStamp)
        {
            TransferOptions transferOptions = new TransferOptions() { TransferMode = TransferMode.Ascii };

            if (!allowResumeSupport)
                transferOptions.ResumeSupport = new TransferResumeSupport() { State = TransferResumeSupportState.Off };

            transferOptions.PreserveTimestamp = allowResumeSupport;

            TransferOperationResult result = Session.PutFiles(sourceFilePath, destinationFilePath + "//" + Path.GetFileName(sourceFilePath), false, transferOptions);

            result.Check();

            if (result.Transfers.Count() != 1)
                return false;
            else
                return true;
        }


        public void CloseSession()
        {
            Session.Close();
        }

    }
}
