using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;


namespace Newsletter
{
    public class SendEmail
    {

        public void SendingEmail(String email)
        {

            try
            {
                //SMTP Configuration
                SmtpClient SmtpServer = new SmtpClient("mail.registrado.com.br");
                SmtpServer.Port = 587;
                SmtpServer.Credentials =
                new System.Net.NetworkCredential("registrado@registrado.com.br", "registrado");
                SmtpServer.EnableSsl = false;
                
                //Create Mail
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("registrado@registrado.com.br");
                mail.To.Add(email);
                mail.Subject = "Assunto";
                mail.IsBodyHtml = true;
                mail.Body = "HTML do corpo aqui";
                 Attachment data = new Attachment(@"arquivo", MediaTypeNames.Application.Octet);
                 mail.Attachments.Add(data);
                SmtpServer.Send(mail);
                               
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadKey();
            }



        }
    }
}
