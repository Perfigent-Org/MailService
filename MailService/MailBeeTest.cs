using MailBee.ImapMail;
using MailService.MailBee;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MailService
{
    public class MailBeeTest
    {
        private readonly Service _mailService;

        IList<long> emailIds = new List<long>();
        IList<string> folderNames = new List<string>();

        public MailBeeTest(Service mailService)
        {
            _mailService = mailService;
        }

        private static MailBeeTest Instance = null;
        public static MailBeeTest GetInstance(Service mailService)
        {
            if (Instance == null)
                Instance = new MailBeeTest(mailService);
            return Instance;
        }

        private async Task GetFoldersAsync()
        {
            Console.WriteLine("\n\n Get all folders");

            var folders = await _mailService.GetFoldersAsync();

            foreach (Folder folder in folders)
            {
                Console.WriteLine($"Folder name: {folder.Name}");

                if (!folder.Name.Contains(_mailService.InboxFolderName))
                    folderNames.Add(folder.Name);
            }

            Console.WriteLine($"Total folders: {folders.Count}");
        }

        private async Task CreateFolderAsync()
        {
            Console.WriteLine("\n\n Create new folder");

            var isDone = await _mailService.CreateFolderAsync("TestMailBeeFolder");

            if (isDone)
            {
                Console.WriteLine("\n\n Folder (TestMailBeeFolder) created successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to create new folder (TestMailBeeFolder), Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task RenameFolderAsync()
        {
            Console.WriteLine("\n\n Rename folder");

            var isDone = await _mailService.RenameFolderAsync("TestMailBeeFolder", "RenameTestMailBee");

            if (isDone)
            {
                Console.WriteLine("\n\n Folder rename from (TestMailBeeFolder) To (RenameTestMailBee) successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to rename the folder (TestMailBeeFolder), Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task DeleteFolderAsync()
        {
            Console.WriteLine("\n\n Delete folder");

            var isDone = await _mailService.DeleteFolderAsync("RenameTestMailBee");

            if (isDone)
            {
                Console.WriteLine("\n\n Folder (RenameTestMailBee) deleted successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to delete the folder (RenameTestMailBee), Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task GetUidsAsync()
        {
            Console.WriteLine("\n\n Get all email id's");

            emailIds = await _mailService.GetUidsAsync();

            foreach (var id in emailIds)
            {
                Console.WriteLine($"Email id: {id}");
            }

            Console.WriteLine($"Total emails: {emailIds.Count}");
        }

        private async Task GetEmailsAsync()
        {
            Console.WriteLine("\n\n Get top 10 emails basic information");

            var emails = await _mailService.GetEmailsAsync(emailIds.Take(10).ToArray());

            foreach (Envelope mail in emails)
            {
                Console.WriteLine($"Email id: {mail.Uid}");
                Console.WriteLine($"Email subject: {mail.Subject}");
            }

            Console.WriteLine($"Total emails: {emails.Count}");
        }

        private async Task SearchEmailsAsync()
        {
            Console.WriteLine($"\n\n Search emails from ({folderNames[4]}) Folder");

            var emailUids = await _mailService.SearchEmailsAsync("ALL", folderNames.FirstOrDefault());

            foreach (var id in emailUids)
            {
                Console.WriteLine($"Email id: {id}");
            }

            Console.WriteLine($"Total email found: {emailUids.Count}");
        }

        private async Task GetEmailAsync()
        {
            Console.WriteLine("\n\n Get email full information");

            var email = await _mailService.GetEmailAsync(emailIds.FirstOrDefault());
            if (email != null)
            {
                Console.WriteLine($"Email id: {email.MessageID}");
                Console.WriteLine($"Email from: {email.From}");
                Console.WriteLine($"Email to: {email.To}");
                Console.WriteLine($"Email subject: {email.Subject}");
                Console.WriteLine($"Email bodyPlainText: {email.BodyPlainText}");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to get email from mail Id: {emailIds.FirstOrDefault()}, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task MarkEmailsAsReadAsync()
        {
            Console.WriteLine("\n\n Mark emails as read");

            var isDone = await _mailService.MarkEmailsAsReadAsync(emailIds.Take(5).ToArray());

            if (isDone)
            {
                Console.WriteLine("\n\n Mark emails as read successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to mark emails as read, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task MarkEmailsAsUnreadAsync()
        {
            Console.WriteLine("\n\n Mark emails as unread");

            var isDone = await _mailService.MarkEmailsAsUnreadAsync(emailIds.Take(5).ToArray());

            if (isDone)
            {
                Console.WriteLine("\n\n Mark emails as unread successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to mark emails as unread, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task FlagEmailsAsync()
        {
            Console.WriteLine("\n\n Flag emails");

            var isDone = await _mailService.FlagEmailsAsync(emailIds.Take(2).ToArray());

            if (isDone)
            {
                Console.WriteLine("\n\n Flag emails successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to flag emails, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task UnflagEmailsAsync()
        {
            Console.WriteLine("\n\n Unflag emails");

            var isDone = await _mailService.UnflagEmailsAsync(emailIds.Take(2).ToArray());

            if (isDone)
            {
                Console.WriteLine("\n\n Unflag emails successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to Unflag emails, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task SendEmailAsync()
        {
            Console.WriteLine($"\n\n Sending email to { _mailService.Email }");

            var isDone = await _mailService.SendEmailAsync("Test Email From MailBee", "This is MailBee Test Email For Testing", new string[] { _mailService.Email });

            if (isDone)
            {
                Console.WriteLine("\n\n Email send successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to send email, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task SendEmailWithAttachmentsAsync()
        {
            Console.WriteLine($"\n\n Sending email to { _mailService.Email }");

            string[] attachments = new string[]
                {
                    @"C:\Users\user\Desktop\image1.jpg",
                    @"C:\Users\user\Desktop\new 2.txt",
                    @"C:\Users\user\Desktop\new tables.xlsx",
                    //@"C:\Users\user\Downloads\Git-2.39.2-64-bit.exe",
                    //@"C:\Users\user\Downloads\MailBeeNetObjects.msi",
                };

            var isDone = await _mailService.SendEmailWithAttachmentsAsync("Test Email From MailBee", "This is MailBee Test Email For Testing", new string[] { _mailService.Email }, attachments);

            if (isDone)
            {
                Console.WriteLine("\n\n Email send successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to send email, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task ReplayEmailAsync()
        {
            Console.WriteLine($"\n\n Replay email");

            var email = await _mailService.GetEmailAsync(emailIds.FirstOrDefault());

            Console.WriteLine($"\n\n Replaying email to { email.ReplyTo }, {email.From}");

            var isDone = await _mailService.ReplayEmailAsync(email, "This is MailBee Test Replay.");

            if (isDone)
            {
                Console.WriteLine("\n\n Email replay successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to replay email, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task ForwardEmailAsync()
        {
            Console.WriteLine($"\n\n Forward email");

            var email = await _mailService.GetEmailAsync(emailIds.FirstOrDefault());

            Console.WriteLine($"\n\n Forwarding email to { _mailService.Email }");

            var isDone = await _mailService.ForwardEmailAsync(email, new string[] { _mailService.Email });

            if (isDone)
            {
                Console.WriteLine("\n\n Email forward successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to forward email, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task ArchiveEmailsAsync()
        {
            Console.WriteLine("\n\n Archive emails");

            var isDone = await _mailService.ArchiveEmailsAsync(emailIds.Take(1).ToArray());

            if (isDone)
            {
                Console.WriteLine("\n\n Archive emails successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to archive emails, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task MoveEmailsAsync()
        {
            Console.WriteLine("\n\n Move emails");
            var toFolderName = folderNames[4];

            var isDone = await _mailService.MoveEmailsAsync(emailIds.Take(1).ToArray(), toFolderName);

            if (isDone)
            {
                Console.WriteLine($"\n\n Move emails form {_mailService.InboxFolderName} to {toFolderName} successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to move emails form {_mailService.InboxFolderName} to {toFolderName}, Please check logfile. {_mailService.LogFileName}");
            }
        }

        private async Task DeleteEmailsAsync()
        {
            Console.WriteLine("\n\n Delete emails");

            var isDone = await _mailService.DeleteEmailsAsync(emailIds.Take(1).ToArray());

            if (isDone)
            {
                Console.WriteLine($"\n\n Delete email successfully.");
            }
            else
            {
                Console.WriteLine($"\n\n Not able to delete the email, Please check logfile. {_mailService.LogFileName}");
            }
        }

        public async Task Run()
        {
            await GetFoldersAsync();
            await CreateFolderAsync();
            await RenameFolderAsync();
            await DeleteFolderAsync();
            await GetUidsAsync();
            await GetEmailsAsync();
            await SearchEmailsAsync();
            await SendEmailWithAttachmentsAsync();
            await GetEmailAsync();
            await MarkEmailsAsReadAsync();
            await MarkEmailsAsUnreadAsync();
            await FlagEmailsAsync();
            await UnflagEmailsAsync();
            await SendEmailAsync();
            await GetUidsAsync();
            await ReplayEmailAsync();
            await ForwardEmailAsync();
            await ArchiveEmailsAsync();
            await MoveEmailsAsync();
            await DeleteEmailsAsync();
        }
    }
}
