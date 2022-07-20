using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GmailAPI.APIHelper;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;

namespace GmailAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("-------System Start----");
            Console.WriteLine("Sending Email....");
            SendingEmail.SendEmail();
            Console.WriteLine("Press any Key to read email....");
            Console.ReadKey();
            Console.WriteLine("Lets Read Email......");
            ReadingEmail();
        }
        public static void ReadingEmail()
        {
            try
            {
                List<Gmail> MailLists = GetAllEmails(Convert.ToString(ConfigurationManager.AppSettings["HostAddress"]));
                Console.WriteLine("--------------Following emails has been read--------------");
                for (int i = 0; i < MailLists.Count; i++)
                {
                    Console.WriteLine("From Email :" + MailLists[i].From + "To :" + MailLists[i].To + "MailDateTime :" + MailLists[i].MailDateTime);
                }
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
            }
        }
        public static List<Gmail> GetAllEmails(string HostEmailAddress)
        {
            try
            {
                GmailService GmailService = GmailAPIHelper.GetService();
                List<Gmail> EmailList = new List<Gmail>();
                UsersResource.MessagesResource.ListRequest ListRequest = GmailService.Users.Messages.List(HostEmailAddress);
                ListRequest.LabelIds = "INBOX";
                ListRequest.IncludeSpamTrash = false;
                ListRequest.Q = "is:unread"; //ONLY FOR UNDREAD EMAIL'S...

                //Get All Unread Emails
                ListMessagesResponse ListResponse = ListRequest.Execute();

                if (ListResponse != null && ListResponse.Messages != null)
                {

                    //loop each email and get the fields which we want 
                    foreach (Message Msg in ListResponse.Messages)
                    {
                        // message marked as read after read email
                        GmailAPIHelper.MsgMarkAsRead(HostEmailAddress, Msg.Id);

                        UsersResource.MessagesResource.GetRequest Message = GmailService.Users.Messages.Get(HostEmailAddress, Msg.Id);
                        Console.WriteLine("\n-----------------READING NEW MAIL----------------------");


                        //Make another request for that email id
                        Message MsgContent = Message.Execute();

                        if (MsgContent != null)
                        {
                            string FromAddress = string.Empty;
                            string Date = string.Empty;
                            string Subject = string.Empty;
                            string MailBody = string.Empty;
                            string ReadableText = string.Empty;

                            //LOOP THROUGH THE HEADERS AND GET THE FIELDS WE NEED (SUBJECT, MAIL)
                            // Loop the headers and get the fields we need (subject)
                            foreach (var MessageParts in MsgContent.Payload.Headers)
                            {
                                if (MessageParts.Name == "From")
                                {
                                    FromAddress = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Date")
                                {
                                    Date = MessageParts.Value;
                                }
                                else if (MessageParts.Name == "Subject")
                                {

                                    Subject = MessageParts.Value;
                                }
                            }
                            //read mail body and get attachments
                            List<string> FileName = new List<string>();
                            if (Subject.Contains("PatientReport_"))
                            {
                                Console.WriteLine("\n-----------------READING CARETEK EMAIL----------------------");
                                FileName = GmailAPIHelper.GetAttachments(HostEmailAddress, Msg.Id, Convert.ToString(ConfigurationManager.AppSettings["GmailAttach"]));
                              
                                //READ MAIL BODY-------------------------------------------------------------------------------------
                                MailBody = string.Empty;
                                if (MsgContent.Payload.Parts == null && MsgContent.Payload.Body != null)
                                {
                                    MailBody = MsgContent.Payload.Body.Data;
                                }
                                else
                                {
                                    MailBody = GmailAPIHelper.MsgNestedParts(MsgContent.Payload.Parts);
                                }

                                //BASE64 TO READABLE TEXT--------------------------------------------------------------------------------
                                ReadableText = string.Empty;
                                ReadableText = GmailAPIHelper.Base64Decode(MailBody);

                                Console.WriteLine("STEP-4: Identifying & Configure Mails.");
                                // add email into list which we have read 
                                if (!string.IsNullOrEmpty(ReadableText))
                                {
                                    Gmail GMail = new Gmail();
                                    GMail.From = FromAddress;
                                    GMail.Body = ReadableText;
                                    GMail.MailDateTime = Convert.ToDateTime(Date);
                                    EmailList.Add(GMail);
                                }
                            }
                        }
                    }
                }
                return EmailList;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);

                return null;
            }
        }
    }
}
