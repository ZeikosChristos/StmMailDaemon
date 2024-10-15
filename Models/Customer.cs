using MailKit.Net.Smtp;
using MimeKit;
using Softone;
using System;
using System.Management;
using System.Threading.Tasks;

namespace StmMailDaemon.Models
{
    public class Customer
    {
        public int Trdr { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string EmailAcc { get; set; }
        public double Balance { get; set; }
        public string StatementPath { get; set; }


        public void GetStatement()
        {
            try
            {
                if (GlobalVariables.ReportType == ReportType.ObjectForm)
                {
                    GetObjectForm();
                }
                else
                {
                    GetReport();
                }

            }
            catch (Exception ex)
            {
                LogWriter.WriteToLog($"{Code}: {ex.Message}", false);
            }
        }

        private void GetObjectForm()
        {
            using var customer = GlobalVariables.xSupport.CreateModule("CUSTOMER");

            customer.LocateData(Trdr);

            var tempPath = $"{GlobalVariables.TempDir}\\{Guid.NewGuid()}.pdf";

            var result = customer.PrintForm(GlobalVariables.ObjectForm, "PDF file", tempPath);

            if (result == "OK")
            {
                StatementPath = tempPath;
            }
        }

        private void GetReport()
        {
            using var customer = GlobalVariables.xSupport.CreateModule("CUSTOMER");

            var tempPath = $"{GlobalVariables.TempDir}\\{Guid.NewGuid()}.pdf";

            var obj = customer.Exec("CODE:SysRequest.executeReport", new object[]
            {
                "CUST_STM",

                GlobalVariables.ReportList,

                $"QUESTIONS.FCODE={Code}&QUESTIONS.TCODE={Code}",

                tempPath,

                GlobalVariables.ReportTemplate
            });

            if (obj.ToString() == tempPath)
            {
                StatementPath = tempPath;
            }
        }

        public void SendStatement()
        {
            try
            {
                if (string.IsNullOrEmpty(StatementPath))
                {
                    return;
                }

                using var customer = GlobalVariables.xSupport.CreateModule("CUSTOMER");

                var mailTo = string.IsNullOrEmpty(EmailAcc) ? Email : EmailAcc;

                var result = customer.Exec("CODE:SysRequest.doSendMail3", new object[]
                {
                    mailTo,

                    string.Empty,

                    string.Empty,

                    GlobalVariables.MailSubject,

                    string.Empty,

                    GlobalVariables.MailHtmlBody,

                    StatementPath,

                    GlobalVariables.MailName,

                    GlobalVariables.MailAccount
                });

                if ((bool)result)
                {
                    LogWriter.WriteToLog($"Mail sent to {Name}", false);
                }
                else
                {
                    LogWriter.WriteToLog($"Αποτυχία αποστολής email σε {Name}", false);
                }
            }
            catch (Exception ex)
            {
                LogWriter.WriteToLog($"{Code}: {ex.Message}", false);
            }
        }
    }
}
