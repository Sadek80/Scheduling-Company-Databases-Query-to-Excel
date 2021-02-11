using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;

namespace test_all_features_2
{
    class SendingEmail
    {
        private static string fromEmail;
        private static string fromEmailPassword;
        private static string toEmail;
        private static string mailSubject;
        private static string mailBodyText;
        private static string excelFilePath;


        public static void setEmailData(string from, string password, string to, string subject, string body, string path)
        {
            fromEmail = from;
            fromEmailPassword = password;
            toEmail = to;
            mailSubject = subject;
            mailBodyText = body;
            excelFilePath = path;
        }


        public static void sendEmail()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient smtp = new SmtpClient("smtp.gmail.com");
                mail.From = new MailAddress(fromEmail);
                mail.To.Add(toEmail);
                mail.Subject = mailSubject;
                mail.Body = mailBodyText;

                System.Net.Mail.Attachment attachment;
                attachment = new Attachment(excelFilePath);
                mail.Attachments.Add(attachment);

                smtp.Port = 587;
                smtp.Credentials = new System.Net.NetworkCredential(fromEmail, fromEmailPassword);
                smtp.EnableSsl = true;
                smtp.Send(mail);



                Console.WriteLine("Email has been successfuly sent!!");
                Console.WriteLine();
                Console.WriteLine();
            }
            catch (Exception e)
            {
                Console.WriteLine("in the send email: "+ e.Message);
            }
        }


    }
}
