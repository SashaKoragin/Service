using System;
using System.IO;
using System.Text.RegularExpressions;
using EfDatabase.Inventory.MailLogicLotus;
using LotusLibrary.MailSender;
using MimeKit;

namespace LibraryOutlook.SubscribeOutlook
{
   public class OutlookAutoSmtp
   {

       public MailSender Mail { get; set; }
        /// <summary>
        /// Отправка писем абоненту внешней почты 
        /// </summary>
        /// <param name="parameters">Параметры конфигурации</param>
        public void SendSmtpMessage(ConfigFile.ConfigFile parameters)
        {
            try
            {
                Mail = new MailSender();
                ZipAttachments zipAttach = new ZipAttachments();
                MailLogicLotus mailSave = new MailLogicLotus();
                var dbSend = Mail.SendMailOut(parameters.PathSaveArchive);
                foreach (var mailLotusOutlookOut in dbSend)
                {
                    //Сначала манипуляции с файлами и архивами а потом отправка
                    var builder = new BodyBuilder() {TextBody = mailLotusOutlookOut.Body};
                    if (mailLotusOutlookOut.FullPathListFile != null)
                    {
                        foreach (var fullFileName in mailLotusOutlookOut.FullPathListFile.Split(';'))
                        {
                            builder.Attachments.Add(fullFileName);
                        }
                        var nameFile = DateTime.Today.ToString("dd.MM.yyyy_HH.mm.ss") + ".zip";
                        var fullPathZip = Path.Combine(parameters.PathSaveArchive, nameFile);
                        mailLotusOutlookOut.FileMailZip = zipAttach.StartZipArchiveOut(mailLotusOutlookOut.FullPathListFile.Split(';'),fullPathZip);
                        mailLotusOutlookOut.NameFileZip = nameFile;
                    }
                    //Проверка почты
                    if (!IsValid(mailLotusOutlookOut.MailAdressOut))
                    {
                        var user = new[] {mailLotusOutlookOut.MailAdressIn};
                        Mail.SendMailAutoOutput(user, $"Ошибка отправки письма на адрес {mailLotusOutlookOut.MailAdressOut} !!!", $"Ошибка почты {mailLotusOutlookOut.MailAdressOut} !!!");
                        mailLotusOutlookOut.ErrorMail = $"Почтовый ящик не прошел проверку отправка не возможна!";
                    }
                    else
                    {
                        MimeMessage mailToClient = new MimeMessage();
                        mailToClient.To.Add(new MailboxAddress(mailLotusOutlookOut.MailAdressOut));
                        mailToClient.Subject = string.IsNullOrWhiteSpace(mailLotusOutlookOut.SubjectMail) ? "" : mailLotusOutlookOut.SubjectMail;
                        mailToClient.From.Add(new MailboxAddress(mailLotusOutlookOut.MailAdressIn, parameters.LoginR7751));
                        mailToClient.Body = builder.ToMessageBody();
                        try
                        {
                            using (var smtp = new MailKit.Net.Smtp.SmtpClient())
                            {
                                smtp.CheckCertificateRevocation = false;
                                smtp.Connect(parameters.Pop3Address, 465, true);
                                smtp.Authenticate(parameters.LoginR7751, parameters.PasswordR7751);
                                smtp.Send(mailToClient);
                                smtp.Disconnect(true);
                                mailLotusOutlookOut.ErrorMail = $"Письмо отправлено адресату {mailLotusOutlookOut.MailAdressOut}";
                            }
                        }
                        catch (Exception ex)
                        {
                            Loggers.Log4NetLogger.Error(ex);
                            mailLotusOutlookOut.ErrorMail = $"Письмо не отправлено адресату возникли ошибки во время отправки";
                        }
                    }
                    mailSave.AddModelMailOut(mailLotusOutlookOut);
                }
                //Очистить временную папку с файлами
                zipAttach.DropAllFileToPath(parameters.PathSaveArchive);
                Loggers.Log4NetLogger.Info(new Exception("Отправка почты внешнем абонентам закончена!"));
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            finally
            {
                Mail.Dispose();
            }
        }

        /// <summary>
        /// Проверка Valid mail User Send
        /// </summary>
        /// <param name="emailAddress">Проверка адреса почты</param>
        /// <returns></returns>
        private bool IsValid(string emailAddress)
        {
            return Regex.IsMatch(emailAddress, @"^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$");
        }
   }
}
