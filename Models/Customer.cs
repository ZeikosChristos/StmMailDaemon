using System;
using System.Collections.Generic;
using System.Text;

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
        public List<string> ObjectForms { get; set; } = new List<string>();
        public List<string> Reports { get; set; } = new List<string>();


        public void GetStatements()
        {
            try
            {
                foreach (var objectForm in GlobalVariables.ObjectForms)
                {
                    if (objectForm.RunDay == 0)
                    {
                        GetObjectForm(objectForm);
                    }
                    else if (objectForm.RunDay >= 31)
                    {
                        if (DateTime.Now.Day == DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month))
                        {
                            GetObjectForm(objectForm);
                        }
                    }
                    else if (objectForm.RunDay == DateTime.Now.Day)
                    {
                        GetObjectForm(objectForm);
                    }
                }

                foreach (var report in GlobalVariables.Reports)
                {
                    if (report.RunDay == 0)
                    {
                        GetReport(report);
                    }
                    else if (report.RunDay >= 31)
                    {
                        if (DateTime.Now.Day == DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month))
                        {
                            GetReport(report);
                        }
                    }
                    else if (report.RunDay == DateTime.Now.Day)
                    {
                        GetReport(report);
                    }
                }
            }
            catch (Exception ex)
            {
                LogWriter.WriteToLog($"{Code}: {ex.Message}", false);
            }
        }

        private void GetObjectForm(ObjectForm objectForm)
        {
            var tempPath = $"{GlobalVariables.TempDir}\\{objectForm.FileName}";

            using var customer = GlobalVariables.xSupport.CreateModule("CUSTOMER");

            customer.LocateData(Trdr);

            var result = customer.PrintForm(objectForm.FormCode, "PDF file", tempPath);

            if (result == "OK")
            {
                ObjectForms.Add(tempPath);
            }
        }

        private void GetReport(Report report)
        {
            var tempPath = $"{GlobalVariables.TempDir}\\{report.FileName}";

            var stockObj = GlobalVariables.xSupport.GetStockObj("SysRequest", true);

            var result = GlobalVariables.xSupport.CallPublished(stockObj, "executeReport", new object[]
            {
                report.ReportObj,

                report.ReportList,

                $"QUESTIONS.FCODE={Code}&QUESTIONS.TCODE={Code}",

                tempPath,

                report.Template
            });

            if (result.ToString() == tempPath)
            {
                Reports.Add(tempPath);
            }
        }

        public void SendStatements()
        {
            try
            {
                if (ObjectForms.Count == 0 && Reports.Count == 0)
                {
                    LogWriter.WriteToLog($"Δεν βρέθηκαν αρχεία για τον πελάτη {Name}", false);

                    return;
                }

                var stringBuilder = new StringBuilder();

                foreach (var objectForm in ObjectForms)
                {
                    stringBuilder.Append($"{objectForm};");
                }

                foreach (var report in Reports)
                {
                    stringBuilder.Append($"{report};");
                }

                var mailTo = string.IsNullOrEmpty(EmailAcc) ? Email : EmailAcc;

                var stockObj = GlobalVariables.xSupport.GetStockObj("SysRequest", true);

                var result = GlobalVariables.xSupport.CallPublished(stockObj, "doSendMail3", new object[]
                {
                    mailTo,

                    string.Empty,

                    string.Empty,

                    GlobalVariables.MailSubject,

                    string.Empty,

                    GlobalVariables.MailHtmlBody,

                    stringBuilder.ToString(0, stringBuilder.Length - 1),

                    GlobalVariables.MailName,

                    GlobalVariables.MailAccount
                });

                if ((bool)result)
                {
                    LogWriter.WriteToLog($"Στάλθηκε email σε {Name}", false);
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
