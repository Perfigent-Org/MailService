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
                            mailService = new Service("466636461440-nob44m3k4c9jr1n3fdtfr4s0vogoh3p4.apps.googleusercontent.com", "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", serverType);
                            break;

                        case ServerType.Office365:
                            mailService = new Service("17b9156a-feec-4c72-ab42-acc6c6d5590a", "Okw8Q~R2IgKW5MZ3P.bcektjRuGCAnOy2M5uBcTZ", serverType);
                            break;

                        case ServerType.Outlook:
                            mailService = new Service("17b9156a-feec-4c72-ab42-acc6c6d5590a", "Okw8Q~R2IgKW5MZ3P.bcektjRuGCAnOy2M5uBcTZ", serverType);
                            break;

                        default:
                            throw new Exception("Server type must be Google, Office365 or Outlook");
                    }
                    break;

                case ServiceType.ImapSmtp:

                    switch (serverType)
                    {
                        case ServerType.Google:
                            mailService = new Service(serverType, "oauthmailbeetest@gmail.com", "fecgqvhghtzovurk");
                            break;

                        case ServerType.Office365:
                            mailService = new Service(serverType, "testoauthmicroauth@outlook.com", "OAuth@MailBee.com");
                            break;

                        case ServerType.Outlook:
                            mailService = new Service(serverType, "testoauthmicroauth@outlook.com", "OAuth@MailBee.com");
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
