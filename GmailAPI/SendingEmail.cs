using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace GmailAPI
{
    public static class SendingEmail
    {
        private static Random random = new Random();
        public static void SendEmail()
        {
            try
            {
                System.IO.MemoryStream ms = new System.IO.MemoryStream();
                System.IO.StreamWriter writer = new System.IO.StreamWriter(ms);
                string str1 = RandomString(50);
                Random rand = new Random();
                //Name: [Some Random Text of length between 3- 20] space [Some Random Text of length between 3- 20].   
                writer.WriteLine("Name :" + str1.Substring(3, 20).Substring(rand.Next(3, 20)) + " " + str1.Substring(3, 20).Substring(rand.Next(3, 20)));
                //Age: Some Random number between 20 and 100.
                writer.WriteLine("Age :" + rand.Next(20, 100).ToString());
                //The system will generate a new file using Memory Stream with name format "MMDDYYYY_HHMMSS.txt"
                string FileNameString = DateTime.Now.ToString("MMddyyyy_HHmmss");
                writer.Flush();
                //writer.Dispose();
                //ms.Position = 0;

                System.Net.Mime.ContentType ct = new System.Net.Mime.ContentType(System.Net.Mime.MediaTypeNames.Text.Plain);
                System.Net.Mail.Attachment attach = new System.Net.Mail.Attachment(ms, ct);
                attach.ContentDisposition.FileName = FileNameString;
                //setup email and send through smtp client
                SmtpClient(attach, FileNameString);
                ms.Close();
                Console.WriteLine("---------Mail Sent-------");            
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Sending Email :" + ex.Message);
            }
        }
        public static void SmtpClient(System.Net.Mail.Attachment attacment, string fileName)
        {
            MailMessage mailMessage = new MailMessage();

            mailMessage.From = new MailAddress("zeeshanilyas002@gmail.com");
            mailMessage.To.Add("zeeshanilyas001@gmail.com");
            mailMessage.Subject = "PatientReport_" + fileName;
            //mailMessage.Body = "Please Find Attachment";
            mailMessage.Attachments.Add(attacment);
            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.EnableSsl = true;
            client.Credentials = new System.Net.NetworkCredential("zeeshanilyas002@gmail.com", "");
            client.Host = "smtp.gmail.com";
            client.Port = 587;

            try
            {
                client.Send(mailMessage);
            }
            catch (SmtpException ex)
            {
                Console.WriteLine("Error while Sending Email :" + ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while Sending Email :" + ex.Message);
            }
        }
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
