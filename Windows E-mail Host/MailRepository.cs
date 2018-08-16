using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

using ActiveUp.Net.Mail;
using System;

namespace Windows_E_mail_Host
{
    public class MailRepository
    {
        private Imap4Client client;

        public MailRepository(string mailServer, int port, bool ssl, string login, string password)
        {
            if (ssl)
                Client.ConnectSsl(mailServer, port);
            else
                Client.Connect(mailServer, port);
            Client.Login(login, password);
        }

        public int[] GetMailIndexes(string mailBox, string searchPhrase)
        {
            Mailbox mails = Client.SelectMailbox(mailBox);
            int[] messagesIndexes = mails.Search(searchPhrase);
            return messagesIndexes;
        }

        public Message GetEmail(string mailBox, int i)
        {
            Mailbox mails = Client.SelectMailbox(mailBox);
            Message message = mails.Fetch.MessageObject(i); ;

            return message;
        }

        public void DelateMessage(string mailBox, int index)
        {
            client.SelectMailbox(mailBox).DeleteMessage(index, true);
        }


        ////////////////////////////////////////////////////////////


        public IEnumerable<Message> GetAllMails(string mailBox)
        {
            return GetMails(mailBox, "ALL").Cast<Message>();
        }

        public IEnumerable<Message> GetUnreadMails(string mailBox)
        {
            return GetMails(mailBox, "UNSEEN").Cast<Message>();
        }

        protected Imap4Client Client
        {
            get { return client ?? (client = new Imap4Client()); }
        }

        private MessageCollection GetMails(string mailBox, string searchPhrase)
        {
            Mailbox mails = Client.SelectMailbox(mailBox);
            MessageCollection messages = mails.SearchParse(searchPhrase);
            return messages;
        }
    }
}
