using MailService.MailBee;
using System;

namespace MailService
{
    public class Program
    {
        public static void Main(string[] args)
        {
            ConnectMailBeeService(ServiceType.OAuth20, ServerType.Google);
            ConnectMailBeeService(ServiceType.OAuth20, ServerType.Outlook);
            ConnectMailBeeService(ServiceType.OAuth20, ServerType.Office365);

            ConnectMailBeeService(ServiceType.ImapSmtp, ServerType.Google);
            ConnectMailBeeService(ServiceType.ImapSmtp, ServerType.Outlook);
            ConnectMailBeeService(ServiceType.ImapSmtp, ServerType.Office365);

            Console.ReadLine();
        }

        private static async void ConnectMailBeeService(ServiceType serviceType, ServerType serverType)
        {
            try
            {
                Service mailService = GetMailService(serviceType, serverType);

                Console.WriteLine($"Connecting {serverType} Mail Service...");

                var isConnected = await mailService.Connect();

                if (isConnected) Console.WriteLine($"{serverType} mail service is ready to use...");

                await MailBeeTest.GetInstance(mailService).Run();

                if (isConnected) await mailService.Disconnect();

                Console.WriteLine($"{serverType} mail service is now disconnected...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static Service GetMailService(ServiceType serviceType, ServerType serverType)
        {
            Service mailService;
            switch (serviceType)
            {
                case ServiceType.OAuth20:

                    switch (serverType)
                    {
                        case ServerType.Google:
                            mailService = new Service("Google-ClientId", "Gmail-ClientSecret", serverType);
                            break;

                        case ServerType.Office365:
                            mailService = new Service("Office365-ClientId", "Office365-ClientSecret", serverType);
                            break;

                        case ServerType.Outlook:
                            mailService = new Service("Outlook-ClientId", "Outlook-ClientSecret", serverType);
                            break;

                        default:
                            throw new Exception("Server type must be Google, Office365 or Outlook");
                    }
                    break;

                case ServiceType.ImapSmtp:

                    switch (serverType)
                    {
                        case ServerType.Google:
                            mailService = new Service(serverType, "AnyAccountName@gmail.com", "AnyAccountPassword");
                            break;

                        case ServerType.Office365:
                            mailService = new Service(serverType, "AnyAccountName@outlook.com", "AnyAccountPassword");
                            break;

                        case ServerType.Outlook:
                            mailService = new Service(serverType, "AnyAccountName@hotmail.com", "AnyAccountPassword");
                            break;

                        default:
                            throw new Exception("Server type must be Google, Office365 or Outlook");
                    }
                    break;

                default:
                    throw new Exception("Service type must be OAuth20 or ImapSmtp");
            }

            return mailService;
        }
    }
}
