using MailBee;
using MailBee.ImapMail;
using MailBee.SmtpMail;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;

namespace MailService.MailBee
{
    public class ServiceBase
    {
        protected readonly string _imapServerName;
        protected readonly string _smtpServerName;
        protected readonly ServerType _serverType;

        protected string _clientId;
        protected string _clientSecret;
        protected string _userEmail;
        protected string _password;
        protected string _accessToken;

        protected readonly HashSet<string> BLACKLISTED_EXTENSIONS = new HashSet<string>
        {
            ".ade", ".adp", ".apk", ".appx", ".appxbundle", ".bat", ".cab", ".chm", ".cmd", ".com", ".cpl", ".dll", ".dmg",
            ".ex", ".ex_", ".exe", ".hta", ".ins", ".isp", ".iso", ".jar", ".js", ".jse", ".lib", ".lnk", ".mde", ".msc",
            ".msi", ".msix", ".msixbundle", ".msp", ".mst", ".nsh", ".pif", ".ps1", ".scr", ".sct", ".shb", ".sys", ".vb",
            ".vbe", ".vbs", ".vxd", ".wsc", ".wsf", ".wsh"
        };

        protected readonly long MAX_ATTACHMENT_SIZE = 19000000; //around 19 MB to account for MIME encoding

        private bool IsFrom0Auth => string.IsNullOrWhiteSpace(_password);

        public string Email => _userEmail;

        public string InboxFolderName { get; set; }
        public string ArchiveFolderName { get; set; }
        public string LogFileName { get; private set; }

        public ServiceBase(ServerType type)
        {
            _serverType = type;

            switch (_serverType)
            {
                case ServerType.Google:
                    _imapServerName = "imap.gmail.com";
                    _smtpServerName = "smtp.gmail.com";
                    break;

                case ServerType.Office365:
                    _imapServerName = "outlook.office365.com";
                    _smtpServerName = "smtp.office365.com";
                    break;

                case ServerType.Outlook:
                    _imapServerName = "imap-mail.outlook.com";
                    _smtpServerName = "smtp-mail.outlook.com";
                    break;
            }

            ArchiveFolderName = "Archive";
            InboxFolderName = "INBOX";
            Global.LicenseKey = "MN120-B078788D785278D0784EC39D5943-4E15";
            LogFileName = @"C:\MailServiceLogs\";
        }

        public async Task<bool> Connect()
        {
            if (IsFrom0Auth)
            {
                switch (_serverType)
                {
                    case ServerType.Google:
                        {
                            return ValidateConnection(await OAuthGoogle.OAuth20.Login(_clientId, _clientSecret));
                        }
                    case ServerType.Office365:
                        {
                            return ValidateConnection(await OAuthOffice365.OAuth20.Login(_clientId));
                        }
                    case ServerType.Outlook:
                        {
                            return ValidateConnection(await OAuthOutlook.OAuth20.Login(_clientId));
                        }
                }
            }

            return ValidateConnection();
        }

        public async Task Disconnect()
        {
            if (IsFrom0Auth)
            {
                switch (_serverType)
                {
                    case ServerType.Google:
                        await OAuthGoogle.OAuth20.Logout(_clientId); break;

                    case ServerType.Office365:
                        await OAuthOffice365.OAuth20.Logout(_clientId); break;

                    case ServerType.Outlook:
                        await OAuthOutlook.OAuth20.Logout(_clientId); break;
                }

                _accessToken = string.Empty;
            }
        }

        protected bool Login(IComponent component)
        {
            AddLogFile(component);

            if (component is Imap imap)
            {
                imap.Connect(_imapServerName);

                if (IsFrom0Auth)
                {
                    string xoauthKey = OAuth2.GetXOAuthKeyStatic(_userEmail, _accessToken);
                    return imap.Login(null, xoauthKey, AuthenticationMethods.SaslOAuth2, AuthenticationOptions.None, null);
                }
                else
                {
                    return imap.Login(_userEmail, _password);
                }
            }
            else if (component is Smtp smtp)
            {
                if (IsFrom0Auth)
                {
                    string xoauthKey = OAuth2.GetXOAuthKeyStatic(_userEmail, _accessToken);
                    smtp.SmtpServers.Add(_smtpServerName, null, xoauthKey, AuthenticationMethods.SaslOAuth2).Port = 587;
                }
                else
                {
                    smtp.SmtpServers.Add(_smtpServerName, _userEmail, _password, AuthenticationMethods.SaslLogin | AuthenticationMethods.SaslPlain).Port = 587;
                }
                return smtp.Connect();
            }
            else
            {
                throw new Exception("The component is not match any type of servers (IMAP or SMTP)");
            }
        }

        private bool ValidateConnection(KeyValuePair<string, string>? userInfo = null)
        {
            if (userInfo.HasValue)
            {
                _userEmail = userInfo.Value.Key;
                _accessToken = userInfo.Value.Value;

                return !string.IsNullOrWhiteSpace(_userEmail) && !string.IsNullOrWhiteSpace(_accessToken);
            }

            return !string.IsNullOrWhiteSpace(_userEmail) && !string.IsNullOrWhiteSpace(_password);
        }

        private void AddLogFile(IComponent component)
        {
            if (component is Imap imap)
            {
                imap.Log.Enabled = true;
                imap.Log.Filename = LogFilePath("Imap");
                imap.Log.HidePasswords = false;
                imap.Log.Clear();
            }
            else if (component is Smtp smtp)
            {
                smtp.Log.Enabled = true;
                smtp.Log.Filename = LogFilePath("Smtp");
                smtp.Log.HidePasswords = false;
                smtp.Log.Clear();
            }
            else
            {
                throw new Exception("The component is not match any type of servers (IMAP or SMTP)");
            }
        }

        private string LogFilePath(string fileType)
        {
            string logFolder = $@"C:\MailServiceLogs\{fileType}_Logs\{_serverType}\{_userEmail}";

            if (!System.IO.Directory.Exists(logFolder))
            {
                System.IO.Directory.CreateDirectory(logFolder);
            }

            LogFileName = $@"{logFolder}\log{DateTime.Now:MMddyyyyhhmmssffftt}.txt";
            return LogFileName;
        }
    }
}