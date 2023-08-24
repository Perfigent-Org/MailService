using MailService.MailBee;
using System;

namespace MailService
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //MailBeeGamilServiceWith0Auth20();

            //MailBeeGamilServiceWithIMAPAndSMTP();

            //MailBeeOffice365ServiceWith0Auth20();

            MailBeeOutlookServiceWith0Auth20();

            Console.ReadLine();
        }

        private async static void MailBeeGamilServiceWith0Auth20()
        {
            try
            {
                var mailService = new Service(ServerType.Google);

                var isConnected = await mailService.Connect();

                if (isConnected) Console.WriteLine("Google mail service is ready to use...");

                await MailBeeTest.GetInstance(mailService).Run();

                if (isConnected) await mailService.Disconnect();

                Console.WriteLine("Google mail service is now disconnected...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private async static void MailBeeGamilServiceWithIMAPAndSMTP()
        {
            try
            {
                var mailService = new Service(ServerType.Google, "chothani.hitesh@gmail.com", "oqdxgsljifsrdtbj");

                var isConnected = await mailService.Connect();

                if (isConnected) Console.WriteLine("Google mail service is ready to use...");

                await MailBeeTest.GetInstance(mailService).Run();

                if (isConnected) await mailService.Disconnect();

                Console.WriteLine("Google mail service is now disconnected...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private async static void MailBeeOffice365ServiceWith0Auth20()
        {
            try
            {
                var mailService = new Service(ServerType.Office365);

                var isConnected = await mailService.Connect();

                if (isConnected) Console.WriteLine("Office365 mail service is ready to use...");

                await MailBeeTest.GetInstance(mailService).Run();

                if (isConnected) await mailService.Disconnect();

                Console.WriteLine("Office365 mail service is now disconnected...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private async static void MailBeeOutlookServiceWith0Auth20()
        {
            try
            {
                var mailService = new Service(ServerType.Outlook);

                var isConnected = await mailService.Connect();

                if (isConnected) Console.WriteLine("Outlook mail service is ready to use...");

                await MailBeeTest.GetInstance(mailService).Run();

                if (isConnected) await mailService.Disconnect();

                Console.WriteLine("Outlook mail service is now disconnected...");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
