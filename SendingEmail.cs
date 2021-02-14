﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Mail;
using System.Timers;

namespace test_all_features_2
{

    class SendingEmail
    {
        private string fromEmail = "mohamedsadk80@gmail.com";
        private string fromEmailPassword = "123221133456554466";
        private List<string> toEmails;
        private string mailSubject;
        private string mailBodyText;
        private string excelFilePath;



        public SendingEmail(List<string> toEmaiilsList , string subject, string body, string path)
        {
            toEmails = toEmaiilsList;
            mailSubject = subject;
            mailBodyText = body;
            excelFilePath = path;
        }

        /// <summary>
        /// Send the email in a certain time
        /// </summary>
        /// <param name="timeToGo">The time that the email will be sent at</param>
        public void sendingEmailInTime(TimeSpan timeToGo)
        {
            // using the Sceduler class and specify the Time and the Task
            Scheduler.Instance.ScheduleTask(timeToGo, () =>
            {
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine("The schedule starts now");
                Console.ResetColor();
                sendEmailNow();
            });

        }

        public void sendEmailNow()
        {
            foreach (string email in toEmails)
            {
                sendEmail(HandlingLocalDb.getDepartmentEmail(email));
            }
        }

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
