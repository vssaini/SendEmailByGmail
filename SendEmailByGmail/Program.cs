using System;

namespace SendEmailByGmail
{
    internal class Program
    {
        private static void Main()
        {
            // ShowLabels();

            SendEmail();

            Console.Read();
        }

        private static void SendEmail()
        {
            Console.WriteLine("Send email:");
            Console.WriteLine("Preparing email message");
            var msg = Gmailer.GetMessage("Hi! This is an test email using Gmail API for DS.", "jyotee.gal@gmail.com", "Jyoti",
                "vs00saini@gmail.com", "Vikram", "Test Email from DS", "Administrator", 3);

            Console.WriteLine("Sending email");
            var gmailMsg = Gmailer.SendMessage("me", msg);

            Console.WriteLine(gmailMsg == null ? "Email sending failed." : "Email was sent successfully!");
        }

        private static void ShowLabels()
        {
            var labels = Gmailer.GetLabels();
            Console.WriteLine("Labels:");
            if (labels != null && labels.Count > 0)
            {
                foreach (var labelItem in labels)
                {
                    Console.WriteLine("{0}", labelItem.Name);
                }
            }
            else
            {
                Console.WriteLine("No labels found.");
            }
        }
    }
}