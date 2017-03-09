using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Mail;
using System.IO;
using System.Net.Mime;
using System.Data;
using System.Configuration;
using System.Security.Cryptography;

namespace ExchangeServiceUtility
{
    class Program
    {
        private static Dictionary<string, LinkedResource> Images;
        static void Main(string[] args)
        {
            //IntialiseService();

            Images = new Dictionary<string, LinkedResource>(StringComparer.OrdinalIgnoreCase);
            LinkedResource WorkFlowFailuresBanner = new LinkedResource(ConfigurationManager.AppSettings["WorkFlowFailuresBanner"]);
            Images.Add("WorkFlowFailuresBanner", WorkFlowFailuresBanner);
            LinkedResource MailboxUpdateAlertBanner = new LinkedResource(ConfigurationManager.AppSettings["MailboxUpdateAlertBanner"]);
            Images.Add("MailboxUpdateAlertBanner", WorkFlowFailuresBanner);
            LinkedResource MicrosoftLogoFooter = new LinkedResource(ConfigurationManager.AppSettings["MicrosoftLogoFooter"]);
            Images.Add("MicrosoftLogoFooter", MicrosoftLogoFooter);

            DataTable t = CreateDataTable();

            MailMessage mail1 = CreateMailMessage(t, "mailboxupdatealert", "Test Mail update", "v-aygu@microsoft.com");
            SendMessage(mail1);

            MailMessage mail2 = CreateMailMessage(t, "workflowfailures", "Test Mail update", "v-aygu@microsoft.com");
            SendMessage(mail2);
        }

        public static void IntialiseService()
        {
            ExchangeService service = new ExchangeService();   //time zone info, specific exchange server version

            string serviceAccountUser = "ccrmappr@microsoft.com";
            string serviceAccountpassword = "R3dApp1!($T1on*";

            //string serviceAccountUser = "fadaappn@microsoft.com";
            //string serviceAccountpassword = "#ppnM!m0spr0d#~";
            service.Credentials = new WebCredentials(serviceAccountUser, serviceAccountpassword); //here we can add domain

            List<string> emails = new List<string>()
            {
                //"eadis@microsoft.com",                "Dmtrtl@microsoft.com",                "Dmtsvcs@microsoft.com",                "mscsupp@microsoft.com",                "disnews@microsoft.com",                "emapoc@microsoft.com",
                "mioentry@microsoft.com",
                "rrlatam@microsoft.com",
                "rrpoe@microsoft.com",
                "xdkar@microsoft.com",
                "xboxar@microsoft.com",
                "PREMESC@microsoft.com"
            };

            foreach (var emailId in emails)
            {
                service.AutodiscoverUrl(emailId, ValidateUrl);  // this will set URL for the particular account 
                Console.WriteLine(service.Url);
                RuleCollection rules = service.GetInboxRules(emailId);
                Console.WriteLine("Success ! for " + emailId + " rules:" + rules.Count);
                Console.WriteLine("**************************************************************");
            }
        }

        public static bool ValidateUrl(string redirectionUrl)
        {
            Console.WriteLine(string.Format("Validating URL: {0}", redirectionUrl));
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. 
            // First, make sure that th eredirection URL is using HTTPS to encrypt the authentication credentials. 
            bool isHttps = redirectionUri.Scheme.ToUpper() == "HTTPS";

            // Next, make sure that the redirect is going through the Outlook.com redirector.
            bool isCorrectHost = redirectionUri.Host == "autodiscover-s.outlook.com" || redirectionUri.Host == "autodiscover.microsoft.com";
            Console.WriteLine(redirectionUri.Host);

            // Finally, make sure that the redirect is to the Autodiscover service.
            bool isCorrectPath = redirectionUri.AbsolutePath == "/autodiscover/autodiscover.xml";

            result = isHttps && isCorrectHost && isCorrectPath;

            if (result)
            {
                Console.WriteLine("Valid");
            }
            else
            {
                Console.WriteLine("Invalid");
            }
            return result;
        }

        public static void SendMessage(MailMessage mail)
        {
            try
            {
                SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SMTPServerName"]);
                SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPort"]);
                string password = Decrypt(ConfigurationManager.AppSettings["MailPassword"]);
                SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["MailUserName"], password);
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public static MailMessage CreateMailMessage(DataTable t, string mailType, string subject, string mailto)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress(ConfigurationManager.AppSettings["MailUserName"]);
            mail.To.Add(mailto);
            mail.Subject = subject;

            EmailTemplateModel model = new EmailTemplateModel(mailType);
            model.Content = TemplateCreator.ConvertDatatableToHtml(t);

            string html = TemplateCreator.CreateEmailTemplate<EmailTemplateModel>(model);

            ContentType mimeType = new System.Net.Mime.ContentType("text/html");
            AlternateView htmlView = AlternateView.CreateAlternateViewFromString(html, mimeType);

            Images[mailType + "Banner"].ContentId = "banner";
            Images[mailType + "Banner"].TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
            htmlView.LinkedResources.Add(Images[mailType + "Banner"]);

            Images["MicrosoftLogoFooter"].ContentId = "footer";
            Images["MicrosoftLogoFooter"].TransferEncoding = System.Net.Mime.TransferEncoding.Base64;
            htmlView.LinkedResources.Add(Images["MicrosoftLogoFooter"]); ;

            mail.AlternateViews.Add(htmlView);
            return mail;
        }
        public static DataTable CreateDataTable()
        {   
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            // Here we add five DataRows.
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            return table;
        }

        private static string Decrypt(string encryptedPassword)
        {
            return Encoding.Unicode.GetString(ProtectedData.Unprotect(Convert.FromBase64String(encryptedPassword), null, DataProtectionScope.LocalMachine));
        }
    }
}
