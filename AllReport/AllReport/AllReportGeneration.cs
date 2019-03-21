using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Configuration;
using System.Net.Mail;
using CommonLibrary.DataDrivenTesting;
using System.Net;
using System.IO;
using Microsoft.Win32;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace AllReport
{
    [TestClass]
    public class AllReportGeneration
    {
        public string Init = ConfigurationManager.AppSettings["Initialization"];
        public string ReportFiles = ConfigurationManager.AppSettings["OtherReports"];
        public string SendingReportConfig = string.Empty;
        private string userName = string.Empty;
        private string password = string.Empty;

        [Description("This Will Cobain all the reports generated batch wise into one file and saved in the folder called AllReports in AutoTestReport based on the date.")]
        [TestMethod]
        public void GenAllReport()
        {
            AutomationTestReportMail();
        }

        [Description("This will send the report to the specified mail the specified in the GlobalElements excel")]
        public void AutomationTestReportMail()
        {
            ExcelDataTable.PopulateInCollection(Init + "\\GlobalElements.xlsx");
            SendingReportConfig = ExcelDataTable.ReadData(1, "ReportSending");
            int Recipients = Convert.ToInt32(ExcelDataTable.ReadData(1, "NumberOfRecipients"));
            string result = string.Empty;
            string to = string.Empty;
            string Emailbody = ExcelDataTable.ReadData(1, "Body").ToString();
            if (SendingReportConfig == "Yes")
            {
                try
                {
                    var fromAddress = new MailAddress(ExcelDataTable.ReadData(1, "MailFrom").ToString(), ExcelDataTable.ReadData(1, "DispayName").ToString());
                    var fromPassword = ExcelDataTable.ReadData(1, "EmailPass").ToString();
                    var toAddress = new MailAddress("deion@vegam.co");
                    string subject = ExcelDataTable.ReadData(1, "Subject").ToString();
                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient
                    {
                        Host = "mail.exclusivehosting.net",
                        Port = 2525,
                        EnableSsl = true,
                        DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network,
                        UseDefaultCredentials = false,
                        Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
                    };
                    ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                    {
                        return true;
                    };
                    using (MailMessage mailMessage = new MailMessage())
                    {
                        mailMessage.From = fromAddress;
                        mailMessage.Subject = subject; 
                        mailMessage.Body = Emailbody;
                        System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(ReportFiles);
                        string[] filepaths = Directory.GetFiles(ReportFiles);
                        foreach (var file in filepaths)
                        {
                            var attachment = new Attachment(file);
                            mailMessage.Attachments.Add(attachment);
                        }
                        Attachment ReportFile;
                        RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Testing");
                        string f= key.GetValue("AutomationReport").ToString();
                        ReportFile = new Attachment(f);
                        mailMessage.Attachments.Add(ReportFile);
                        mailMessage.IsBodyHtml = true;
                        for (int i = 1; i <= Recipients; i++)
                        {
                            try
                            {
                                to = ExcelDataTable.ReadData(i, "MailTo").ToString();
                                mailMessage.To.Add(new MailAddress(to));
                            }
                            catch (Exception e) { }
                        }
                        using (var message = new MailMessage(fromAddress, toAddress)
                        {
                            Subject = subject,
                            Body = Emailbody,
                        })
                        smtp.Send(mailMessage);
                    }
                }
                catch (Exception e) { }
            }
            else
            {
                Console.WriteLine("Sending Batch Report is disabled");
            }
        }

    }
}
