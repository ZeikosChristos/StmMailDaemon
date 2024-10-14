using MailKit.Net.Smtp;
using MimeKit;
using Softone;
using System;
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

            var tempPath = $"{AppDomain.CurrentDomain.BaseDirectory}CusStm\\{Guid.NewGuid()}.pdf";

            var result = customer.PrintForm(GlobalVariables.ObjectForm, "PDF file", tempPath);

            if (result == "OK")
            {
                StatementPath = tempPath;
            }
        }

        private void GetReport()
        {
            using var customer = GlobalVariables.xSupport.CreateModule("CUSTOMER");

            var tempPath = $"{AppDomain.CurrentDomain.BaseDirectory}CusStm\\{Guid.NewGuid()}.pdf";

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

        public async Task SendStatement(SmtpClient smtpClient)
        {
            try
            {
                if (string.IsNullOrEmpty(StatementPath))
                {
                    return;
                }

                var message = new MimeMessage();

                message.From.Add(new MailboxAddress(GlobalVariables.MailName, GlobalVariables.MailAddress));

                if (!string.IsNullOrEmpty(EmailAcc))
                {
                    message.To.Add(new MailboxAddress(EmailAcc, EmailAcc));
                }
                else
                {
                    message.To.Add(new MailboxAddress(Email, Email));
                }

                message.Subject = GlobalVariables.MailSubject;

                var builder = new BodyBuilder()
                {
                    HtmlBody = GlobalVariables.MailHtmlBody
                };

                builder.Attachments.Add(StatementPath);

                message.Body = builder.ToMessageBody();

                await smtpClient.SendAsync(message);

                LogWriter.WriteToLog($"Mail sent to {Name}", false);
            }
            catch (Exception ex)
            {
                LogWriter.WriteToLog($"{Code}: {ex.Message}", false);
            }
        }
    }
}
