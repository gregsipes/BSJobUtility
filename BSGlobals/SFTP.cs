﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinSCP;

namespace BSGlobals
{
    public class SFTP
    {
        private Session Session { get; set; }

        private string Host { get; set; }

        private string UserName { get; set; }

        private string Password { get; set; }

        public SFTP(string host, string userName, string password)
        {
            Host = host;
            UserName = userName;
            Password = password;
        }

        public void OpenSession(string fingerprint, string keyFilePath, string keyPassPhrase)
        {
            SessionOptions sessionOptions = new SessionOptions()
            {
                Protocol = Protocol.Sftp,
                HostName = Host,
                UserName = UserName,
                Password = Password,
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
            TransferOptions transferOptions = new TransferOptions() { TransferMode = TransferMode.Binary, PreserveTimestamp = true };

            if (!allowResumeSupport)
                transferOptions.ResumeSupport = new TransferResumeSupport() { State = TransferResumeSupportState.Off };

            transferOptions.PreserveTimestamp = allowResumeSupport;

            TransferOperationResult result = Session.PutFiles(sourceFilePath, destinationFilePath + "/" + Path.GetFileName(sourceFilePath), false, transferOptions);

            result.Check();

            if (result.Transfers.Count() != 1)
                return false;
            else
                return true;
        }

        public void CreateDirectory(string path)
        {
            Session.CreateDirectory(path);
        }

        public bool CheckIfFileOrDirectoryExists(string path)
        {
            return Session.FileExists(path);
        }



        public void CloseSession()
        {
            Session.Close();
        }

        public List<string> GetFiles(string path, string inputMask)
        {
            //return only the file name
           return Session.EnumerateRemoteFiles(path, inputMask, EnumerationOptions.AllDirectories).ToList().Select(r => r.FullName).ToList();

        }

        public void MoveFile(string path, string destination)
        {
            Session.MoveFile(path, destination);
        }

    }
}
