using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace SendEmailByGmail
{
    internal class Gmailer
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        private static readonly string[] Scopes = { GmailService.Scope.GmailReadonly,GmailService.Scope.GmailSend };
        private const string ApplicationName = "AD Self Password Reset";

        /// <summary>
        /// Get or set the credential file path.
        /// </summary>
        public static string CredFilePath { get; set; }

        private static GmailService _gmailService;

        // Constructor
        static Gmailer()
        {
            string credFilePath;
            InitializeGmailService(out credFilePath);
            CredFilePath = credFilePath;
        }

        /// <summary>
        /// Initialize the GmailService using the credential file and application.
        /// </summary>
        /// <param name="credPath"></param>
        public static void InitializeGmailService(out string credPath)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credPath = Environment.GetFolderPath(
                    Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/gmail-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                // Write to console
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // If token expired
            if (credential.Token.IsExpired(credential.Flow.Clock))
            {
                Console.WriteLine("The access token has expired, refreshing it");
                if (credential.RefreshTokenAsync(CancellationToken.None).Result)
                {
                    Console.WriteLine("The access token is now refreshed");
                }
                else
                {
                    Console.WriteLine("The access token has expired but we can't refresh it :(");
                    return;
                }
            }
            else
            {
                Console.WriteLine("The access token is OK, continue");
            }

            // Create Gmail API service.
            _gmailService = new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        /// <summary>
        /// Set message as per Gmail standards.
        /// </summary>
        /// <param name="template">Content of email template</param>
        /// <param name="senderEmail">Email address of sender(From)</param>
        /// <param name="senderName">Name of sender</param>
        /// <param name="recipientEmail">Email address of recipient (To)</param>
        /// <param name="recipientName">Name of recipient</param>
        /// <param name="emailSubject">Subject of email</param>
        /// <param name="samAccountName"></param>
        /// <param name="daysToExpire"></param>
        /// <returns>Return message in Gmail format.</returns>
        public static Message GetMessage(string template, string senderEmail, string senderName, string recipientEmail, string recipientName, string emailSubject, string samAccountName = null, int daysToExpire = 0)
        {
            // Create a new mail message
            var mail = new MailMessage
            {
                From = new MailAddress(senderEmail, senderName)
            };

            // Set To and Subject            
            mail.To.Add(new MailAddress(recipientEmail, recipientName));
            mail.Subject = emailSubject;

            // Get the template HTML content and replace placeholder
            template = template.Replace("{DISPLAY_NAME}", recipientName);
            template = template.Replace("{ACCOUNT_NAME}", samAccountName);
            template = template.Replace("{DAYS_TO_EXPIRE}", daysToExpire.ToString());

            // Create plain text view of the message
            var plainText = CleanHtmlString(template);
            var plainView = AlternateView.CreateAlternateViewFromString(plainText, Encoding.UTF8, "text/plain");

            // Create html view
            var htmlView = AlternateView.CreateAlternateViewFromString(template, Encoding.UTF8, "text/html");

            // Add two views to message
            mail.AlternateViews.Add(plainView);
            mail.AlternateViews.Add(htmlView);

            //var mail = new MailMessage
            //{
            //    Subject = subject,
            //    Body = bodyHtml,
            //    From = new MailAddress(senderEmail),
            //    IsBodyHtml = true
            //};

            //// Mail to users in list
            //foreach (var add in userEmailList.Where(add => !string.IsNullOrEmpty(add)))
            //{
            //    mail.To.Add(new MailAddress(add));
            //}

            //foreach (var path in attachments)
            //{
            //    //var bytes = File.ReadAllBytes(path);
            //    //string mimeType = MimeMapping.GetMimeMapping(path);
            //    var attachment = new Attachment(path);//bytes, mimeType, Path.GetFileName(path), true);
            //    mail.Attachments.Add(attachment);
            //}

            var mimeMessage = MimeKit.MimeMessage.CreateFromMailMessage(mail);
            var gmailMsg = new Message { Raw = BaseUrl64Encode(mimeMessage.ToString()) };
            return gmailMsg;

            //var result = gmailService.Users.Messages.Send(message, "me").Execute();
        }

        /// <summary>
        /// Send an email from the user's mailbox to its recipient.
        /// </summary>
        /// <param name="userId">User's email address. The special value "me"
        /// can be used to indicate the authenticated user.</param>
        /// <param name="email">Email to be sent.</param>
        public static Message SendMessage(string userId, Message email)
        {
            return SendMessage(_gmailService, userId, email);
        }

        /// <summary>
        /// Send an email from the user's mailbox to its recipient.
        /// </summary>
        /// <param name="service">Gmail API service instance.</param>
        /// <param name="userId">User's email address. The special value "me"
        /// can be used to indicate the authenticated user.</param>
        /// <param name="email">Email to be sent.</param>
        public static Message SendMessage(GmailService service, string userId, Message email)
        {
            try
            {
                // await service.Mylibrary.Bookshelves.List().ExecuteAsync();
                return service.Users.Messages.Send(email, userId).Execute();
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred: " + e.Message);
            }

            return null;
        }

        /// <summary>
        /// Get list of Labels as per the authenticated account.
        /// </summary>
        /// <returns>Return list of labels.</returns>
        public static IList<Label> GetLabels()
        {
            // Define parameters of request.
            var request = _gmailService.Users.Labels.List("me");

            // List labels.
            var labels = request.Execute().Labels;
            return labels;
        }

        /// <summary>
        /// Removes all HTML tags from an HTML string
        /// </summary>
        /// <param name="html">The HTML string</param>
        /// <returns>The plain text string</returns>
        public static string CleanHtmlString(string html)
        {
            // Replace multiple spaces with one space
            html = Regex.Replace(html, @"\s+", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Replace BR tags with line breaks
            html = Regex.Replace(html, @"<br[^>/]*/?>", "\r\n", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Replace </p> tags with two line breaks
            html = Regex.Replace(html, @"</p>", "\r\n\r\n", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Replace </div> tags with two line breaks
            html = Regex.Replace(html, @"</div>", "\r\n\r\n", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Replace </h1>,</h2>,.. tags with two line breaks
            html = Regex.Replace(html, @"</h\d>", "\r\n\r\n", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // remove ALL html tags
            html = Regex.Replace(html, @"<[^>]*>", "", RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Decode HTML
            html = System.Web.HttpUtility.HtmlDecode(html);

            // Trim lines one by one
            string cleanText = string.Empty;

            using (StringReader sr = new StringReader(html))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                    cleanText += line.Trim() + "\r\n";

                sr.Close();
            }

            cleanText = cleanText.Trim();

            return cleanText;
        }

        /// <summary>
        /// Encode the passed value into Base64 string.
        /// </summary>
        /// <param name="text">The text that need to encode.</param>
        /// <returns>Return encoded string.</returns>
        public static string BaseUrl64Encode(string text)
        {
            var bytes = Encoding.UTF8.GetBytes(text);

            var base64String = Convert.ToBase64String(bytes)
                .Replace('+', '-')
                .Replace('/', '_')
                .Replace("=", "");

            return base64String;
        }
    }
}
