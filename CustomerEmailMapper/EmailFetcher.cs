using System;
using System.Collections.Generic;
using System.Net.Mail;
using S22.Imap;
using System.Text.RegularExpressions;

namespace CustomerEmailMapper
{
    class EmailFetcher
    {
        public string username;

        public string password;

        public UIForm form;

        public string folderName;

        public IEnumerable<uint> uids;


        public EmailFetcher(string username, string password, UIForm form, string folderName)
        {
            this.username = username;
            this.password = password;
            this.form = form;
            this.folderName = folderName;
        }

        public List<CustomerInfoTemplate> FetchEmailInfo()
        {
            try
            {
                using (ImapClient ic = new ImapClient("imap.gmail.com", 993, username, password, AuthMethod.Login, true))
                {
                    ic.DefaultMailbox = "INBOX";
                    List<CustomerInfoTemplate> allInfo = new List<CustomerInfoTemplate>();
                    uids = ic.Search(SearchCondition.Subject("NEW BUYER").And(SearchCondition.Undeleted()).And(SearchCondition.Unseen()));
                    IEnumerable<MailMessage> messages = ic.GetMessages(uids);
                    foreach (MailMessage m in messages)
                    {
                        Console.WriteLine(m.Body);
                        CustomerInfoTemplate template = new CustomerInfoTemplate();

                        var rx = new Regex(@"(?<=Email:\r\n).*(?=<)", RegexOptions.IgnoreCase);
                        Match match = rx.Match(m.Body);
                        template.email = match.Value;

                        rx = new Regex(@"(?<=Phone:\r\n).*(?=\r\n)");
                        match = rx.Match(m.Body);
                        template.phone = match.Value;

                        rx = new Regex(@"(?<=View\sBuyer\sProfile<).*(?=>)");
                        match = rx.Match(m.Body);
                        template.profileUrl = match.Value;

                        rx = new Regex(@"(?<=Registration\r\nfrom\s).*(?=\!\r\n)");
                        match = rx.Match(m.Body);
                        template.name = match.Value;

                        allInfo.Add(template);
                    }
                    
                    form.WriteStatus("Success!");
                    ic.Dispose();

                    return allInfo;
                }
            }
            catch (Exception ex)
            {
                if (ex is InvalidCredentialsException)
                {
                    form.WriteStatus("Username or password incorrect");
                    throw ex;
                }
                else
                {
                    form.WriteStatus("Failure " + ex.Message);
                    throw ex;
                }
            }
        }

        public void MoveMessages()
        {
            ImapClient ic = new ImapClient("imap.gmail.com", 993, username, password, AuthMethod.Login, true);
            foreach (uint id in uids)
            {
                ic.MoveMessage(id, folderName);
            }
        }

        public void MarkAsUnseen()
        {
            ImapClient ic = new ImapClient("imap.gmail.com", 993, username, password, AuthMethod.Login, true);
            foreach (uint id in uids)
            {
                ic.RemoveMessageFlags(id, "INBOX", MessageFlag.Seen);
            }
        }
    }
}
