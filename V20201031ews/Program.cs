using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;

namespace V20201031ews
{
    class Program
    {
        // 2020-10-31 Visual Studio Community 2015 Update 3 Console Application. 
        // Using Microsoft Exchange Web Services (EWS) Managed API 2.2 to connect to Exchange 365 Online
        static void Main(string[] args)
        {
            Console.WriteLine("{0} Start.", DateTime.Now);
            // dotnet add package Microsoft.Exchange.WebServices --version 2.2.0
            // https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications?redirectedfrom=MSDN
            //string fromEmail = "";
            //string fromPwd = "";
            string fromEmail = ConfigurationManager.AppSettings["fromEmail"];
            string fromPwd = ConfigurationManager.AppSettings["fromPwd"];
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
            try
            {
                if (ConfigurationManager.AppSettings["traceFlag"].ToLower()=="true")
                {
                    service.TraceEnabled = true;
                    service.TraceFlags = TraceFlags.All;
                }
                service.Credentials = new WebCredentials(fromEmail, fromPwd);
                //Don't use it. 401 Unauthorize error
                //service.UseDefaultCredentials = true;
                service.WebProxy = null;
                service.PreAuthenticate = true;
                service.AutodiscoverUrl(fromEmail, RedirectionUrlValidationCallback);

                EmailMessage email = new EmailMessage(service);
                email.ToRecipients.Add("abc@example.org");
                email.Subject = "Hello World";
                email.Body = new MessageBody("This is the test email I've sent using the Microsoft Exchange Web Services (EWS) Managed API.");
                email.Send();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Error: {1}", DateTime.Now, e.Message);
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

    }
}
