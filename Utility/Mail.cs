using System.Net.Mail;  // MailMessage


namespace InstructorBriefcaseExtractor.Utility
{
    public class Mail
    {
        public Mail()
        { }

        public static void EmailDeveloper(string EmailTo, string subject, string body)
        {
            MailMessage myMessage;
            try
            {
                // Use exchange server to send email
                myMessage = new MailMessage(Properties.Settings.Default.SMTPUser,
                                            EmailTo,
                                            "IBE Usage From: " + subject,
                                            body)
                {
                    BodyEncoding = System.Text.Encoding.ASCII
                };

                //Send the message.
                SmtpClient client = new SmtpClient(Properties.Settings.Default.SMTPServer, Properties.Settings.Default.SMTPPort)
                {
                    EnableSsl = true,
                    UseDefaultCredentials = false,

                    // Add credentials if the SMTP server requires them.
                    Credentials = new System.Net.NetworkCredential(Properties.Settings.Default.SMTPUser, Properties.Settings.Default.SMTPPassword), 

                    Timeout = 10000,
                    DeliveryMethod = SmtpDeliveryMethod.Network
                };

                client.Send(myMessage);
            }
            catch 
            {
                // silently fail
            }

        }

    }
}
