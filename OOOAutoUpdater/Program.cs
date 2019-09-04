using AzureResourceReport.Models;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Management.ResourceManager.Fluent;
using Microsoft.Azure.Management.ResourceManager.Fluent.Authentication;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Configuration;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;

namespace UpdateOOO
{
    class Program
    {
        private readonly IConfigurationRoot _configuration;

        public Program()
        {
            var config = new ConfigurationBuilder()
            .AddEnvironmentVariables();
            _configuration = config.Build();
        }
        static void Main(string[] args)
        {
            TimeWindow duration;
            Boolean SetOOF;
            TimeSpan starttime, endtime;
            DateTime dt = DateTime.UtcNow.Date;
            switch (dt.DayOfWeek)
            {
                case DayOfWeek.Monday:
                    SetOOF = false;
                    starttime = new TimeSpan(14, 30, 0);
                    endtime = new TimeSpan(8, 15, 0);
                    duration = getDuration(starttime, endtime, 1);
                    break;
                case DayOfWeek.Tuesday:
                    SetOOF = true;
                    starttime = new TimeSpan(14, 15, 0);
                    endtime = new TimeSpan(9, 0, 0);
                    duration = getDuration(starttime, endtime, 1);
                    break;
                case DayOfWeek.Wednesday:
                    SetOOF = true;
                    starttime = new TimeSpan(17, 0, 0);
                    endtime = new TimeSpan(12, 0, 0);
                    duration = getDuration(starttime, endtime, 1);
                    break;
                case DayOfWeek.Thursday:
                    SetOOF = true;
                    starttime = new TimeSpan(14, 30, 0);
                    endtime = new TimeSpan(8, 15, 0);
                    duration = getDuration(starttime, endtime, 1);
                    break;
                case DayOfWeek.Friday:
                    SetOOF = true;
                    SetOOF = true;
                    starttime = new TimeSpan(14, 15, 0);
                    endtime = new TimeSpan(8, 15, 0);
                    duration = getDuration(starttime, endtime, 4);
                    break;
                default:
                    SetOOF = false;
                    starttime = new TimeSpan(14, 30, 0);
                    endtime = new TimeSpan(8, 15, 0);
                    duration = getDuration(starttime, endtime, 1);
                    break;
            }
            if (SetOOF)
            {
                UpdateOOO(duration);
            }
            // Dont do anything 
        }
        private static void If(bool v)
        {
            throw new NotImplementedException();
        }
        public static void UpdateOOO(TimeWindow duration)
        {
            var email_user = Environment.GetEnvironmentVariable("email_user");
            var email = Environment.GetEnvironmentVariable("email");
            var ews_url = Environment.GetEnvironmentVariable("ews_url");
            string clientId = Environment.GetEnvironmentVariable("akvClientId");
            string clientSecret = Environment.GetEnvironmentVariable("akvClientSecret");
            string tenantId = Environment.GetEnvironmentVariable("akvTenantId");
            string subscriptionId = Environment.GetEnvironmentVariable("akvSubscriptionId");
            string kvURL = Environment.GetEnvironmentVariable("akvName");
            string secretName = Environment.GetEnvironmentVariable("akvSecret");
            AzureCredentials credentials = SdkContext.AzureCredentialsFactory.FromServicePrincipal(clientId, clientSecret, tenantId, AzureEnvironment.AzureGlobalCloud).WithDefaultSubscription(subscriptionId);
            var keyClient = new KeyVaultClient(async (authority, resource, scope) =>
            {
                var adCredential = new ClientCredential(clientId, clientSecret);
                var authenticationContext = new AuthenticationContext(authority, null);
                return (await authenticationContext.AcquireTokenAsync(resource, adCredential)).AccessToken;
            });
            KeyVaultCache keyVaultCache = new KeyVaultCache(kvURL, clientId, clientSecret);
            var cacheSecret = keyVaultCache.GetCachedSecret(secretName);
            string email_pass = cacheSecret.Result;
            // Create the binding.
            ExchangeService service = new ExchangeService();
            service.TraceEnabled = false;
            service.TraceFlags = TraceFlags.All;
            // Set the credentials for the on-premises server.
            service.Credentials = new WebCredentials(email_user, email_pass);
            // Set the URL.
            service.Url = new Uri(ews_url);
            // Return the Out Of Office object that contains OOF state for the user whose credendials were supplied at the console. 
            // This method will result in a call to the Exchange Server.
            OofSettings userOOFSettings = service.GetUserOofSettings(email);
            OofSettings userOOF = new OofSettings();
            // Select the OOF status to be a set time period.
            userOOF.State = OofState.Scheduled;
            // Select the time period to be OOF
            userOOF.Duration = duration;
            // Select the external audience that will receive OOF messages.
            userOOF.ExternalAudience = OofExternalAudience.All;
            // Select the OOF reply for your internal audience.
            userOOF.InternalReply = userOOFSettings.InternalReply;
            // Select the OOF reply for your external audience.
            userOOF.ExternalReply = userOOFSettings.ExternalReply;
            service.SetUserOofSettings(email, userOOF);
            Console.WriteLine("Updated OOF");
            Console.WriteLine("StartDate {0:MM/dd/yy H:mm:ss zzz}", duration.StartTime);
            Console.WriteLine("EndDate {0:MM/dd/yy H:mm:ss zzz}", duration.EndTime);
            return;
        }

        public static TimeWindow getDuration(TimeSpan starttime, TimeSpan endtime, int days)
        {
            TimeWindow duration;            
            DateTime dt = DateTime.Today;
            DateTime startdate, enddate, datetomorrow;
            startdate = dt.Add(starttime);
            datetomorrow = dt.AddDays(days);
            enddate = datetomorrow.Add(endtime);            
            duration = new TimeWindow(startdate, enddate);
            return duration;
        }
    }
}
