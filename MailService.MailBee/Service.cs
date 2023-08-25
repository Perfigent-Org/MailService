using MailBee.ImapMail;
using MailBee.Mime;
using MailBee.SmtpMail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MailService.MailBee
{
    public class Service : ServiceBase, IService
    {
        public Service(ServerType type, string userEmail, string password) : base(type)
        {
            _userEmail = userEmail;
            _password = password;
        }

        public Service(string clientId, string clientSecret, ServerType type) : base(type)
        {
            _clientId = clientId;
            _clientSecret = clientSecret;
        }

        public async Task<IList<long>> GetUidsAsync(string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            using (var imap = new Imap())
            {
                if (Login(imap))
                {
                    imap.SelectFolder(fromFolder);

                    var allUids = await imap.SearchAsync();

                    await imap.DisconnectAsync();

                    return allUids.ToArray().Select(s => s).ToList();
                }

                return new List<long>();
            }
        }

        public async Task<FolderCollection> GetFoldersAsync()
        {
            using (var imap = new Imap())
            {
                if (Login(imap))
                {
                    var folders = await imap.DownloadFoldersAsync();

                    await imap.DisconnectAsync();

                    return folders;
                }

                return null;
            }
        }

        public async Task<bool> CreateFolderAsync(string folderName)
        {
            using (var imap = new Imap())
            {
                var result = false;

                if (Login(imap))
                {
                    result = await imap.CreateFolderAsync(folderName);

                    await imap.DisconnectAsync();
                }

                return result;
            }
        }

        public async Task<bool> RenameFolderAsync(string oldFolderName, string newFolderName)
        {
            using (var imap = new Imap())
            {
                var result = false;

                if (Login(imap))
                {
                    result = await imap.RenameFolderAsync(oldFolderName, newFolderName);

                    await imap.DisconnectAsync();
                }

                return result;
            }
        }

        public async Task<bool> DeleteFolderAsync(string folderName)
        {
            using (var imap = new Imap())
            {
                var result = false;

                if (Login(imap))
                {
                    result = await imap.DeleteFolderAsync(folderName);
                    await imap.DisconnectAsync();
                }

                return result;
            }
        }

        public async Task<EnvelopeCollection> GetEmailsAsync(long[] uids, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            using (var imap = new Imap())
            {
                if (Login(imap))
                {
                    imap.SelectFolder(fromFolder);

                    var uidsString = string.Join(",", uids);

                    var messageHeaders = await imap.DownloadEnvelopesAsync(uidsString, true);

                    await imap.DisconnectAsync();

                    return messageHeaders;
                }

                return null;
            }
        }

        public async Task<IList<long>> SearchEmailsAsync(string searchCondition, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            using (var imap = new Imap())
            {
                if (Login(imap))
                {
                    imap.SelectFolder(fromFolder);

                    var a = await imap.SearchAsync(true, searchCondition, null);

                    var matchingUids = (UidCollection)(await imap.SearchAsync(true, searchCondition, null));

                    await imap.DisconnectAsync();

                    return matchingUids.ToArray().Select(s => s).ToList();
                }

                return new List<long>();
            }
        }

        public async Task<MailMessage> GetEmailAsync(long uid, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            using (var imap = new Imap())
            {
                if (Login(imap))
                {
                    imap.SelectFolder(fromFolder);

                    var message = await imap.DownloadEntireMessageAsync(uid, true);

                    await imap.DisconnectAsync();

                    return message;
                }

                return null;
            }
        }

        public async Task<bool> MarkEmailsAsReadAsync(long[] uids, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            return await SetFlagAsync(uids, fromFolder, SystemMessageFlags.Seen, MessageFlagAction.Add);
        }

        public async Task<bool> MarkEmailsAsUnreadAsync(long[] uids, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            return await SetFlagAsync(uids, fromFolder, SystemMessageFlags.Seen, MessageFlagAction.Remove);
        }

        public async Task<bool> FlagEmailsAsync(long[] uids, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            return await SetFlagAsync(uids, fromFolder, SystemMessageFlags.Flagged, MessageFlagAction.Add);
        }

        public async Task<bool> UnflagEmailsAsync(long[] uids, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            return await SetFlagAsync(uids, fromFolder, SystemMessageFlags.Flagged, MessageFlagAction.Remove);
        }

        public async Task<bool> SendEmailAsync(string subject, string body, string[] to, string[] cc = null, string[] bcc = null, bool isHtml = false)
        {
            using (var smtp = new Smtp())
            {
                var result = false;

                if (Login(smtp))
                {
                    MailMessage message = new MailMessage();

                    message.From.AsString = _userEmail;
                    message.To.AsString = string.Join(",", to);

                    if (cc != null)
                        message.Cc.AsString = string.Join(",", cc);

                    if (bcc != null)
                        message.Bcc.AsString = string.Join(",", bcc);

                    message.Subject = subject;

                    if (isHtml)
                    {
                        message.BodyHtmlText = body;
                    }
                    else
                    {
                        message.BodyPlainText = body;
                    }

                    smtp.Message = message;

                    result = await smtp.SendAsync();

                    await smtp.DisconnectAsync();
                }

                return result;
            }
        }

        /// <summary>
        ///  Gmail: The maximum attachment size is 25 MB.
        ///  Yahoo Mail: The maximum attachment size is 25 MB.
        ///  Outlook/Hotmail: The maximum attachment size is 20 MB (or up to 10 GB with OneDrive integration).
        ///  iCloud Mail: The maximum attachment sie is 20 MB (or up to 5 GB with Mail Drop feature).
        /// </summary>
        public async Task<bool> SendEmailWithAttachmentsAsync(string subject, string body, string[] to, string[] attachmentPaths, string[] cc = null, string[] bcc = null, bool isHtml = false)
        {
            using (var smtp = new Smtp())
            {
                var result = false;

                if (Login(smtp))
                {
                    MailMessage message = new MailMessage();

                    message.From.AsString = _userEmail;
                    message.To.AsString = string.Join(",", to);

                    if (cc != null)
                        message.Cc.AsString = string.Join(",", cc);

                    if (bcc != null)
                        message.Bcc.AsString = string.Join(",", bcc);

                    message.Subject = subject;

                    if (isHtml)
                    {
                        message.BodyHtmlText = body;
                    }
                    else
                    {
                        message.BodyPlainText = body;
                    }

                    long totalAttachmentSize = 0;

                    foreach (var path in attachmentPaths)
                    {
                        FileInfo fileInfo = new FileInfo(path);

                        if (BLACKLISTED_EXTENSIONS.Contains(fileInfo.Extension.ToLower()))
                        {
                            throw new Exception($"File {path} has a blacklisted extension.");
                        }

                        if (totalAttachmentSize + fileInfo.Length <= MAX_ATTACHMENT_SIZE)
                        {
                            message.Attachments.Add(path);
                            totalAttachmentSize += fileInfo.Length;
                        }
                        else
                        {
                            throw new Exception($"Attaching file {path} would exceed the maximum email size.");
                        }
                    }

                    smtp.Message = message;

                    result = await smtp.SendAsync();

                    await smtp.DisconnectAsync();
                }

                return result;
            }
        }

        public async Task<bool> ReplayEmailAsync(MailMessage originalEmail, string replayBody)
        {
            using (var smtp = new Smtp())
            {
                var result = false;

                if (Login(smtp))
                {
                    MailMessage message = new MailMessage();

                    message.From.AsString = _userEmail;
                    message.To.AsString = originalEmail.From;

                    if (!string.IsNullOrWhiteSpace(originalEmail.Cc))
                        message.Cc.AsString = originalEmail.Cc;

                    if (!string.IsNullOrWhiteSpace(originalEmail.Bcc))
                        message.Bcc.AsString = originalEmail.Bcc;

                    message.Subject = $"Re: {originalEmail.Subject}";

                    if (string.IsNullOrWhiteSpace(originalEmail.BodyHtmlText))
                    {
                        message.BodyPlainText = $"{replayBody} \n\n --------------- Original message --------------- \n {originalEmail.BodyPlainText}";
                    }
                    else
                    {
                        message.BodyHtmlText = $"{replayBody} \n\n --------------- Original message --------------- \n {originalEmail.BodyHtmlText}";
                    }

                    message.Headers["In-Replay-To"] = originalEmail.MessageID;

                    smtp.Message = message;

                    result = await smtp.SendAsync();

                    await smtp.DisconnectAsync();
                }

                return result;
            }
        }

        public async Task<bool> ForwardEmailAsync(MailMessage originalEmail, string[] to)
        {
            using (var smtp = new Smtp())
            {
                var result = false;

                if (Login(smtp))
                {
                    MailMessage message = new MailMessage();

                    message.From.AsString = _userEmail;
                    message.To.AsString = string.Join(",", to);

                    message.Subject = $"Fwd: {originalEmail.Subject}";

                    message.Attachments.Add(new Attachment(MimePart.Parse(originalEmail.GetMessageRawData())));

                    smtp.Message = message;

                    result = await smtp.SendAsync();

                    await smtp.DisconnectAsync();
                }

                return result;
            }
        }

        public async Task<bool> ArchiveEmailsAsync(long[] uids, string fromFolder = "")
        {
            if (_serverType == ServerType.Google)
            {
                return await DeleteEmailsAsync(uids, fromFolder);
            }
            else
            {
                return await MoveEmailsAsync(uids, ArchiveFolderName, fromFolder);
            }
        }

        public async Task<bool> MoveEmailsAsync(long[] uids, string toFolder, string fromFolder = "")
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            using (var imap = new Imap())
            {
                var result = false;

                if (Login(imap))
                {
                    var uidsString = string.Join(",", uids);

                    imap.SelectFolder(fromFolder);

                    result = await imap.MoveMessagesAsync(uidsString, true, toFolder);

                    await imap.DisconnectAsync();
                }

                return result;
            }
        }

        public async Task<bool> DeleteEmailsAsync(long[] uids, string fromFolder = "", bool hardDelete = false)
        {
            if (string.IsNullOrWhiteSpace(fromFolder))
                fromFolder = InboxFolderName;

            using (var imap = new Imap())
            {
                var result = false;

                if (Login(imap))
                {
                    imap.SelectFolder(fromFolder);

                    var uidsString = string.Join(",", uids);

                    result = await imap.DeleteMessagesAsync(uidsString, true);

                    if (hardDelete) await imap.ExpungeAsync(uidsString, true);

                    await imap.DisconnectAsync();
                }

                return result;
            }
        }

        private async Task<bool> SetFlagAsync(long[] uids, string fromFolder, SystemMessageFlags flag, MessageFlagAction action)
        {
            using (var imap = new Imap())
            {
                var result = false;

                if (Login(imap))
                {
                    imap.SelectFolder(fromFolder);

                    var uidsString = string.Join(",", uids);

                    result = await imap.SetMessageFlagsAsync(uidsString, true, flag, action);

                    await imap.DisconnectAsync();
                }

                return result;
            }
        }
    }
}
