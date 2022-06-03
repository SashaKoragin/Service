using MailKit;
using MimeKit;
using SmtpClient = MailKit.Net.Smtp.SmtpClient;

namespace LibraryOutlook.SmtpCustomClient
{
   public class SmtpCustomClient : SmtpClient
    {

        public SmtpCustomClient()
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
