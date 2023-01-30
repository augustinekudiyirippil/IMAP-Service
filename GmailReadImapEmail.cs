using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using ActiveUp.Net.Mail;


namespace AKIMAPService
{
    public class MailRepository
    {
        private Imap4Client client;
        private Mailbox mails;

        public MailRepository(string mailServer, int port, bool ssl, string login, string password)
        {
            try
            {
                if (ssl)
                    Client.ConnectSsl(mailServer, port);
                else
                    Client.Connect(mailServer, port);
                Client.Login(login, password);
            }
            catch (Exception e)
            {

                Console.WriteLine(e);
                System.Environment.Exit(-1);
            }

        }

        public IEnumerable<Message> GetAllMails(string mailBox)
        {
            return GetMails(mailBox, "ALL");
        }

        public IEnumerable<Message> GetUnreadMails(string mailBox)
        {
            return GetMails(mailBox, "UNSEEN");

            //return GetMails(mailBox, "ALL");
        }




        // client.delete_message(emailId)

        protected Imap4Client Client
        {
            get { return client ?? (client = new Imap4Client()); }
        }
        public Mailbox Mails
        {
            get
            {
                return mails;
            }

            set { mails = value; }
        }

        private IEnumerable<Message> GetMails(string mailBox, string searchPhrase)
        {
            int count = 1;
            mails = Client.SelectMailbox(mailBox);
            var Messages = mails.SearchParse(searchPhrase).Cast<Message>();
            var messagesAll = mails.SearchParse("All").Cast<Message>();


            foreach (Message Email in Messages)
            {
                foreach (Message EmailAll in messagesAll)
                {
                    if (EmailAll.MessageId == Email.MessageId) Email.Id = count;

                    count++;
                }
                count = 1;
            }
            return Messages;
        }



    }
}
