using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Utility.Email
{
    public static class Utils
    {
        #region Create Mail Message

        public static MailMessage ToMailMessage(Mail mail)
        {
            if (mail == null) throw new ArgumentNullException(nameof(mail));

            return ToMailMessage(mail.From, mail.To, mail.Subject, mail.HtmlBody, mail.TextBody, mail.CC, mail.BCC, mail.ReplyTo, mail.Attach);
        }

        public static MailMessage ToMailMessage(string mailFrom, string mailTo, string mailSubject, string htmlMailBody, string textMailBody, string mailCC, string mailBCC, string mailReplyTo, string attachFiles)
        {
            MailMessage mailMsg = new MailMessage
            {
                From = new MailAddress(mailFrom),
                Subject = mailSubject,
                BodyEncoding = Encoding.UTF8
            };

            AddAddresses(mailMsg.To, mailTo);
            AddAddresses(mailMsg.CC, mailCC);
            AddAddresses(mailMsg.Bcc, mailBCC);
            AddAddresses(mailMsg.Bcc, mailBCC);
            AddAddresses(mailMsg.ReplyToList, mailReplyTo);

            SetBody(mailMsg, htmlMailBody, textMailBody);
            AddAttachments(mailMsg, attachFiles);

            return mailMsg;
        }

        #endregion

        #region Send Mail

        public static SendingResult SendWithCredentials(Mail mail, Smtp server)
        {
            return SendWithCredentials(ToMailMessage(mail), server);
        }

        public static SendingResult SendWithCredentials(Mail mail)
        {
            return SendWithCredentials(ToMailMessage(mail));
        }

        public static SendingResult SendWithCredentials(MailMessage mail, Smtp server)
        {
            if (mail == null) throw new ArgumentNullException(nameof(mail));
            if (server == null) throw new ArgumentNullException(nameof(server));

            SmtpClient client = new SmtpClient(server.RelayServer, server.RelayPort)
            {
                EnableSsl = server.UseSSL,
                Credentials = new NetworkCredential(server.UserName ?? string.Empty, server.Password ?? string.Empty)
            };

            return SendInternal(mail, client);
        }

        public static SendingResult SendWithCredentials(MailMessage mail)
        {
            if (mail == null) throw new ArgumentNullException(nameof(mail));

            SmtpClient client = new SmtpClient();
            return SendInternal(mail, client);
        }

        #endregion

        #region Internal Send Logic

        private static SendingResult SendInternal(MailMessage mail, SmtpClient client)
        {
            SendingResult results = new SendingResult();

            RemoteCertificateValidationCallback originalCallback = ServicePointManager.ServerCertificateValidationCallback;
            ServicePointManager.ServerCertificateValidationCallback = ValidateServerCertificate;

            try
            {
                results.First = new SingleResult().BeginSend();
                client.Send(mail);
                results.First.EndSend();
            }
            catch (SmtpException ex)
            {
                results.First.EndSend(ex);

                if (IsTimeout(ex))
                {
                    try
                    {
                        results.TryAgain = new SingleResult().BeginSend();
                        client.Send(mail);
                        results.TryAgain.EndSend();
                    }
                    catch (SmtpException retryEx)
                    {
                        results.TryAgain.EndSend(retryEx);
                    }
                }
            }
            finally
            {
                ServicePointManager.ServerCertificateValidationCallback = originalCallback;
                mail.Dispose();
                client.Dispose();
            }

            return results;
        }

        private static bool IsTimeout(SmtpException ex)
        {
            return ex.StatusCode == SmtpStatusCode.GeneralFailure || ex.InnerException is TimeoutException;
        }

        #endregion

        #region Helpers

        private static void AddAddresses(MailAddressCollection collection, string addresses)
        {
            if (string.IsNullOrWhiteSpace(addresses)) return;
            foreach (string addr in addresses.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(a => a.Trim()).Where(SafeConversion.IsEmail))
            {
                collection.Add(new MailAddress(addr));
            }
        }

        private static void SetBody(MailMessage mailMsg, string htmlBody, string textBody)
        {
            if (!string.IsNullOrEmpty(htmlBody) && !string.IsNullOrEmpty(textBody))
            {
                mailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(textBody, Encoding.UTF8, "text/plain"));
                mailMsg.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(htmlBody, Encoding.UTF8, "text/html"));
            }
            else if (!string.IsNullOrEmpty(htmlBody))
            {
                mailMsg.Body = htmlBody;
                mailMsg.IsBodyHtml = true;
            }
            else
            {
                mailMsg.Body = textBody ?? string.Empty;
            }
        }

        private static void AddAttachments(MailMessage mailMsg, string attachFiles)
        {
            if (string.IsNullOrWhiteSpace(attachFiles)) return;

            foreach (string filename in attachFiles.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(f => f.Trim()).Where(File.Exists))
            {
                Attachment attachment = new Attachment(filename, MediaTypeNames.Application.Octet);

                ContentDisposition disposition = attachment.ContentDisposition;
                disposition.CreationDate = File.GetCreationTime(filename);
                disposition.ModificationDate = File.GetLastWriteTime(filename);
                disposition.ReadDate = File.GetLastAccessTime(filename);

                mailMsg.Attachments.Add(attachment);
            }
        }

        #endregion

        #region Certificate Validation (保持原行为)

        [System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Assert, Unrestricted = true)]
        public static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }

        #endregion
    }
}
