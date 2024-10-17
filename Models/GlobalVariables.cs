using Softone;
using System;
using System.Collections.Generic;
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
        public static List<ObjectForm> ObjectForms { get; private set; }
        public static List<Report> Reports { get; private set; }
        public static int MailBatchSize { get; private set; }
        public static int MailBatchDelay { get; private set; }
        public static int MailAccount { get; private set; }
        public static string MailName { get; private set; }
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

            TempDir = $"{AppDomain.CurrentDomain.BaseDirectory}CusStm";

            SqlFilter = ConfigurationManager.AppSettings.Get("SqlFilter");

            SqlFilter ??= string.Empty;

            ObjectForms = ObjectForm.GetObjectForms(ConfigurationManager.AppSettings.Get("ObjectForms"));

            Reports = Report.GetReports(ConfigurationManager.AppSettings.Get("Reports"));

            MailBatchSize = int.TryParse(ConfigurationManager.AppSettings.Get("MailBatchSize"), out int mailBatchSize) ? mailBatchSize : 0;

            MailBatchDelay = int.TryParse(ConfigurationManager.AppSettings.Get("MailBatchDelay"), out int mailBatchDelay) ? mailBatchDelay : 0;

            MailAccount = int.TryParse(ConfigurationManager.AppSettings.Get("MailAccount"), out int mailAccount) ? mailAccount : 0;

            MailName = ConfigurationManager.AppSettings.Get("MailName");

            MailName ??= string.Empty;

            MailSubject = ConfigurationManager.AppSettings.Get("MailSubject");

            if (string.IsNullOrEmpty(MailSubject))
            {
                MailSubject = "Αποστολή Καρτέλας";
            }

            MailHtmlBody = ConfigurationManager.AppSettings.Get("MailHtmlBody");

            MailHtmlBody ??= string.Empty;
        }

    }

    public class ObjectForm
    {
        public int RunDay { get; set; }
        public int FormCode { get; set; }
        public string FileName { get; set; }


        public static List<ObjectForm> GetObjectForms(string configString)
        {
            var objectForms = new List<ObjectForm>();

            if (string.IsNullOrEmpty(configString))
            {
                return objectForms;
            }

            var formsArray = configString.Split('|');

            foreach (var formArray in formsArray)
            {
                if (formArray.Length < 3)
                {
                    continue;
                }

                var formProperties = formArray.Split(';');

                var objectForm = new ObjectForm()
                {
                    RunDay = int.TryParse(formProperties[0], out int id) ? id : 0,

                    FormCode = int.TryParse(formProperties[1], out int day) ? day : 0,

                    FileName = formProperties[2] == string.Empty ? $"{Guid.NewGuid()}.pdf" : $"{formProperties[2]}.pdf"
                };

                objectForms.Add(objectForm);
            }

            return objectForms;
        }
    }

    public class Report
    {
        public int RunDay { get; set; }
        public string ReportObj { get; set; }
        public string ReportList { get; set; }
        public string Template { get; set; }
        public string FileName { get; set; }

        public static List<Report> GetReports(string configString)
        {
            var reports = new List<Report>();

            if (string.IsNullOrEmpty(configString))
            {
                return reports;
            }

            var reportsArray = configString.Split('|');

            foreach (var reportArray in reportsArray)
            {
                if (reportArray.Length < 5)
                {
                    continue;
                }

                var reportProperties = reportArray.Split(';');

                var report = new Report()
                {
                    RunDay = int.TryParse(reportProperties[0], out int day) ? day : 0,

                    ReportObj = reportProperties[1] == string.Empty ? "CUST_STM" : reportProperties[1],

                    ReportList = reportProperties[2] == string.Empty ? "Πρότυπο" : reportProperties[2],

                    Template = reportProperties[3],

                    FileName = reportProperties[4] == string.Empty ? $"{Guid.NewGuid()}.pdf" : $"{reportProperties[4]}.pdf"
                };

                reports.Add(report);
            }

            return reports;
        }
    }
}
