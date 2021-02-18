using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Timers;
using System.Threading;
namespace test_all_features_2
{
    /// <summary>
    /// Class for Sending the Emails
    /// </summary>
    class SendingEmail
    {
        // Default Email Data
        private string fromEmail = "mohamedsadk80@gmail.com";
        private string fromEmailPassword = "123221133456554466";

        /// <summary>
        /// List of Department Emails
        /// </summary>
        private List<string> toEmails;

        /// <summary>
        /// Single Department Email
        /// </summary>
        private string toEmail;

        /// <summary>
        /// Mail Subject
        /// </summary>
        private string mailSubject;

        /// <summary>
        /// Mail Body
        /// </summary>
        private string mailBodyText;

        /// <summary>
        /// The Path of the Excel File Name that will be sent to the Departments
        /// </summary>
        private string excelFilePath;

        /// <summary>
        /// Name of the Company
        /// </summary>
        private string companyName;


        /// <summary>
        /// Public Constructor to Setting Up the Default Sending Email Data, Used for Sending the Email Now
        /// </summary>
        /// <param name="toEmaiilsList">List of Departments Emails</param>
        /// <param name="subject">Mail Subject</param>
        /// <param name="body">Mail Body</param>
        /// <param name="path">Excel File Path</param>
        /// <param name="companyName">Company Name</param>
        public SendingEmail(List<string> toEmaiilsList , string subject, string body, string path, string companyName)
        {
            toEmails = toEmaiilsList;
            mailSubject = subject;
            mailBodyText = body;
            excelFilePath = path;
            this.companyName = companyName;
        }


        /// <summary>
        /// Public Constructor for Setting Up the Sending Email Data, Used for the Schedule Tasks
        /// </summary>
        /// <param name="toEmail">Department Email</param>
        /// <param name="mailSubject">Mail Subject</param>
        /// <param name="mailBodyText">Mail Body</param>
        /// <param name="excelFilePath">Excel File Path</param>
        public SendingEmail(string toEmail, string mailSubject, string mailBodyText, string excelFilePath)
        {
            this.toEmail = toEmail;
            this.mailSubject = mailSubject;
            this.mailBodyText = mailBodyText;
            this.excelFilePath = excelFilePath;
        }



        /// <summary>
        /// Send the email in a specific time, used for the Schedule Tasks
        /// </summary>
        public void sendingEmail()
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
                Console.WriteLine("Email has been sent successfuly!!");
                Console.ResetColor();
                Console.WriteLine();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the send email: " + e.Message);
                Console.ResetColor();
            }

        }


        /// <summary>
        /// Sending Email Now, for Every Department
        /// </summary>
        public void sendEmailNow()
        {
            foreach (string email in toEmails)
            {
                sendEmail(HandlingLocalDb.getDepartmentEmail(email, companyName));
            }
        }

        /// <summary>
        /// Sending Email for the Department passed
        /// </summary>
        /// <param name="toEmail">Department Email</param>
        public void sendEmail(string toEmail)
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
                Console.WriteLine("Email has been sent successfuly!!");
                Console.ResetColor();
                Console.WriteLine();
            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("in the send email: "+ e.Message);
                Console.ResetColor();
            }
            }

    }
}
