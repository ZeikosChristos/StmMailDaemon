using Softone;
using StmMailDaemon.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.ServiceProcess;
using System.Threading.Tasks;


namespace StmMailDaemon
{
    public partial class StMailDaemon : ServiceBase
    {
        public StMailDaemon()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            GlobalVariables.Initialize();

            LogWriter.WriteToLog("Service Started");

            Task.Run(() => SendCustomerStatements());
        }

        protected override void OnStop()
        {
            LogWriter.WriteToLog("Service Stopped");
        }

        private async Task SendCustomerStatements()
        {
            try
            {
                XSupport.InitInterop(0, GlobalVariables.XdllPath);

                GlobalVariables.xSupport = XSupport.Login(GlobalVariables.XcoPath, GlobalVariables.S1Username, GlobalVariables.S1Password, GlobalVariables.Company, GlobalVariables.Branch, DateTime.Now);

                var cusDataset = GlobalVariables.xSupport.GetSQLDataSet("select TRDR.TRDR, TRDR.CODE, TRDR.NAME, TRDR.EMAIL, TRDR.EMAILACC, TRDFINDATA.LBAL " +
                                                                        "from TRDR inner join TRDEXTRA on TRDR.TRDR = TRDEXTRA.TRDR inner join TRDFINDATA  on TRDR.COMPANY = TRDFINDATA.COMPANY and TRDR.TRDR = TRDFINDATA.TRDR " +
                                                                        $"where TRDR.COMPANY = :1 and TRDR.SODTYPE = 13 and (TRDR.EMAIL is not null or TRDR.EMAILACC is not null) {GlobalVariables.SqlFilter}",
                                                                        GlobalVariables.Company).CreateDataTable(true);

                if (cusDataset.Rows.Count == 0)
                {
                    return;
                }

                Directory.CreateDirectory(GlobalVariables.TempDir);

                var customers = new List<Customer>();

                foreach (DataRow row in cusDataset.Rows) 
                {
                    var customer = new Customer()
                    {
                        Trdr = Convert.ToInt32(row["TRDR"]),

                        Code = row["CODE"].ToString(),

                        Name = row["NAME"].ToString(),

                        Email = row["EMAIL"].ToString(),

                        EmailAcc = row["EMAILACC"].ToString(),

                        Balance = row["LBAL"] == DBNull.Value ? 0d : Convert.ToDouble(row["LBAL"])
                    };
                    
                    customers.Add(customer);
                }

                var counter = 1;

                foreach (var customer in customers)
                {
                    customer.GetStatements();

                    customer.SendStatements();
                    
                    counter++;

                    if (counter > GlobalVariables.MailBatchSize)
                    {
                        await Task.Delay(GlobalVariables.MailBatchDelay);

                        counter = 1;
                    }
                }
            }
            catch (Exception ex)
            {
                LogWriter.WriteToLog(ex.Message, false);
            }
            finally
            {
                foreach (var file in Directory.GetFiles(GlobalVariables.TempDir))
                {
                    File.Delete(file);
                }

                GlobalVariables.xSupport?.Close();

                GlobalVariables.xSupport?.Dispose();

                Stop();
            }
        }
    }
}
