using MailBee.ImapMail;
using MailBee.Mime;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MailService.MailBee
{
    public interface IService
    {
        Task<FolderCollection> GetFoldersAsync();
        Task<bool> CreateFolderAsync(string folderName);
        Task<bool> RenameFolderAsync(string oldFolderName, string newFolderName);
        Task<bool> DeleteFolderAsync(string folderName);
        Task<IList<long>> GetUidsAsync(string fromFolder);
        Task<EnvelopeCollection> GetEmailsAsync(long[] uids, string fromFolder);
        Task<IList<long>> SearchEmailsAsync(string searchCondition, string fromFolder);
        Task<MailMessage> GetEmailAsync(long uid, string fromFolder);
        Task<bool> MarkEmailsAsReadAsync(long[] uids, string fromFolder);
        Task<bool> MarkEmailsAsUnreadAsync(long[] uids, string fromFolder);
        Task<bool> FlagEmailsAsync(long[] uids, string fromFolder);
        Task<bool> UnflagEmailsAsync(long[] uids, string fromFolder);
        Task<bool> SendEmailAsync(string subject, string body, string[] to, string[] cc = null, string[] bcc = null, bool isHtml = false);
        Task<bool> SendEmailWithAttachmentsAsync(string subject, string body, string[] to, string[] attachmentPaths, string[] cc = null, string[] bcc = null, bool isHtml = false);
        Task<bool> ReplayEmailAsync(MailMessage originalEmail, string replayBody);
        Task<bool> ForwardEmailAsync(MailMessage originalEmail, string[] to);
        Task<bool> ArchiveEmailsAsync(long[] uids, string fromFolder);
        Task<bool> MoveEmailsAsync(long[] uids, string fromFolder, string toFolder);
        Task<bool> DeleteEmailsAsync(long[] uids, string fromFolder, bool hardDelete = false);
    }
}
