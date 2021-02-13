using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;

namespace test_all_features_2
{
    class SendingEmail
    {
        private  static string fromEmail = "mohamedsadk80@gmail.com";
        private  static string fromEmailPassword = "123221133456554466";
        private static List<string> toEmails;
        private static string mailSubject;
        private static string mailBodyText;
        private static string excelFilePath;


        public static void setEmailData(List<string> toEmaiilsList , string subject, string body, string path)
        {
            toEmails = toEmaiilsList;
            mailSubject = subject;
            mailBodyText = body;
            excelFilePath = path;
            
        }


        public static void sendEmailNow()
        {
            foreach (string email in toEmails)
            {
                sendEmail(HandlingLocalDb.getDepartmentEmail(email));
            }
        }

        public static void sendEmail(string toEmail)
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


                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Email has been successfuly sent!!");
                Console.ResetColor();
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
