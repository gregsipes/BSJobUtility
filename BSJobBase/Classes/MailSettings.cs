using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BSJobBase.Classes
{
   public class MailSettings
    {
        public bool UseTLS { get; set; }
        public string Host { get; set; }
        public int Port { get; set; }
        public string User { get; set; }
        public string Password { get; set; }
        public string DefaultSender { get; set; }
        public string DefaultRecipient { get; set; }
    }
}
