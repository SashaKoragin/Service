using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using MailKit;
using MimeKit;
using SmtpClient = MailKit.Net.Smtp.SmtpClient;

namespace LibraryOutlook.SmtpCastomClient
{
   public class SmtpCastomClient : SmtpClient
    {

        public SmtpCastomClient()
        {

        }
        protected override string GetEnvelopeId(MimeMessage message)
        {
            return message.MessageId;
        }
        protected override DeliveryStatusNotification? GetDeliveryStatusNotifications(MimeMessage message, MailboxAddress mailbox)
        {
            return DeliveryStatusNotification.Success|DeliveryStatusNotification.Failure|DeliveryStatusNotification.Delay;
        }

   
    }
}
