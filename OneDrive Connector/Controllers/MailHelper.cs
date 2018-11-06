using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OneDrive_Connector.Controllers
{
    class MailHelper
    {
        public static Message ComposeMail(string subject, string body, List<string> recipients)
        {
            List<Recipient> recipientsList = new List<Recipient>();

            foreach(var recipient in recipients)
            {
                recipientsList.Add(new Recipient { EmailAddress = new EmailAddress { Address = recipient } });
            }

            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = body,
                    ContentType = BodyType.Text,
                },
                Subject = subject,
                ToRecipients = recipientsList,
            };

            return email;
        }
    }
}
