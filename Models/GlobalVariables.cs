using Softone;
using System;
using System.Configuration;


namespace StmMailDaemon.Models
{
    public static class GlobalVariables
    {
        
        public static XSupport xSupport { get; set; }
        public static string XdllPath { get; private set; }
        public static string XcoPath { get; private set; }
        public static int Company { get; private set; }
        public static int Branch { get; private set; }
        public static string S1Username { get; private set; }
        public static string S1Password { get; private set; }
        public static string LogPath { get; private set; }
        public static string TempDir { get; private set; }
        public static string SqlFilter { get; private set; }
        public static int ObjectForm { get; private set; }
        public static int MailBatchSize { get; private set; }
        public static int MailBatchDelay { get; private set; }
        public static string MailServer { get; private set; }
        public static string MailUsername { get; private set; }
        public static string MailPassword { get; private set; }
        public static string MailAddress { get; private set; }
        public static string MailName { get; private set; }
        public static int MailPort { get; private set; }
        public static bool MailSSL { get; private set; }
        public static string MailSubject { get; private set; }
        public static string MailHtmlBody { get; private set; }


        public static void Initialize()
        {
            XdllPath = ConfigurationManager.AppSettings.Get("XdllPath");

            if (XdllPath == string.Empty)
            {
                XdllPath = null;
            }

            XcoPath = ConfigurationManager.AppSettings.Get("XcoPath");

            XcoPath ??= string.Empty;

            Company = int.TryParse(ConfigurationManager.AppSettings.Get("Company"), out int company) ? company : 0;

            Branch = int.TryParse(ConfigurationManager.AppSettings.Get("Branch"), out int branch) ? branch : 0;

            S1Username = ConfigurationManager.AppSettings.Get("S1Username");

            S1Username ??= string.Empty;

            S1Password = ConfigurationManager.AppSettings.Get("S1Password");

            S1Password ??= string.Empty;

            LogPath = ConfigurationManager.AppSettings.Get("LogPath");

            if (string.IsNullOrEmpty(LogPath))
            {
                LogPath = $"{AppDomain.CurrentDomain.BaseDirectory}StmMailLog.txt";
            }

            TempDir = ConfigurationManager.AppSettings.Get("TempDir");

            if (string.IsNullOrEmpty(TempDir))
            {
                TempDir = $"{AppDomain.CurrentDomain.BaseDirectory}CusStm";
            }

            SqlFilter = ConfigurationManager.AppSettings.Get("SqlFilter");

            SqlFilter ??= string.Empty;

            ObjectForm = int.TryParse(ConfigurationManager.AppSettings.Get("ObjectForm"), out int objectForm) ? objectForm : 0;

            MailBatchSize = int.TryParse(ConfigurationManager.AppSettings.Get("MailBatchSize"), out int mailBatchSize) ? mailBatchSize : 0;

            MailBatchDelay = int.TryParse(ConfigurationManager.AppSettings.Get("MailBatchDelay"), out int mailBatchDelay) ? mailBatchDelay : 0;

            MailServer = ConfigurationManager.AppSettings.Get("MailServer");

            MailServer ??= string.Empty;

            MailUsername = ConfigurationManager.AppSettings.Get("MailUsername");

            MailUsername ??= string.Empty;

            MailPassword = ConfigurationManager.AppSettings.Get("MailPassword");

            MailPassword ??= string.Empty;

            MailAddress = ConfigurationManager.AppSettings.Get("MailAddress");

            MailAddress ??= string.Empty;

            MailName = ConfigurationManager.AppSettings.Get("MailName");

            MailName ??= string.Empty;

            MailPort = int.TryParse(ConfigurationManager.AppSettings.Get("MailPort"), out int mailPort) ? mailPort : 0;

            MailSSL = bool.TryParse(ConfigurationManager.AppSettings.Get("MailSSL"), out bool mailSSL) ? mailSSL : false;

            MailSubject = ConfigurationManager.AppSettings.Get("MailSubject");

            if (string.IsNullOrEmpty(MailSubject))
            {
                MailSubject = "Αποστολή Καρτέλας";
            }

            MailHtmlBody = ConfigurationManager.AppSettings.Get("MailHtmlBody");

            MailHtmlBody ??= string.Empty;
        }

    }
}
