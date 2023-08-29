using MailService.MailBee;
using System;
using System.Configuration;

namespace MailService
{
    public class Program
    {
        private static string clientId;
        private static string clientSecret;
        private static string userEmail;
        private static string password;
        private static ServerType serverType;
        private static AuthenticationMode serviceType;

        public static void Main(string[] args)
        {
            serverType = (ServerType)Convert.ToInt32(ConfigurationManager.AppSettings["ServerType"]);

            serviceType = (AuthenticationMode)Convert.ToInt32(ConfigurationManager.AppSettings["AuthenticationMode"]);

            clientId = ConfigurationManager.AppSettings["ClientId"];
            clientSecret = ConfigurationManager.AppSettings["ClientSecret"];

            userEmail = ConfigurationManager.AppSettings["UserEmail"];
            password = ConfigurationManager.AppSettings["Password"];

            ConnectMailBeeService();

            Console.ReadLine();
        }

        private static async void ConnectMailBeeService()
        {
            bool isConnected = false;
            Service mailService = null;
            try
            {
                mailService = GetMailService();

                Console.WriteLine($"Connecting {serverType} Mail Service...");

                isConnected = await mailService.Connect();

                if (isConnected) Console.WriteLine($"{serverType} mail service is ready to use...");

                await MailBeeTest.GetInstance(mailService).Run();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
            finally
            {
                if (isConnected) await mailService.Disconnect();

                Console.WriteLine($"\n\n{serverType} mail service is now disconnected...");
            }
        }

        private static Service GetMailService()
        {
            switch (serviceType)
            {
                case AuthenticationMode.OAuth: return new Service.ServiceBuilder(serverType).WithOAuthCredentials(clientId, clientSecret).Build();

                case AuthenticationMode.UserCredentials: return new Service.ServiceBuilder(serverType).WithUserCredentials(userEmail, password).Build();

                default: throw new Exception("Service type must be OAuth or UserCredentials");
            }
        }
    }
}
