using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
                    var user = new List<string>() {mailLotusOutlookOut.MailAdressIn};
                    var arrayMail = MailArraySubject(mailLotusOutlookOut.MailAdressOut);
                    if (arrayMail.Length > 0)
                    {
                        mailLotusOutlookOut.ErrorMail = $"Письмо отправлено адресатам {string.Join("/", arrayMail)}";
                        foreach (var mail in arrayMail)
                        {
                            MimeMessage mailToClient = new MimeMessage();
                            mailToClient.To.Add(new MailboxAddress(mail));
                            mailToClient.Subject = string.IsNullOrWhiteSpace(mailLotusOutlookOut.SubjectMail) ? "" : mailLotusOutlookOut.SubjectMail;
                            mailToClient.From.Add(new MailboxAddress(mailLotusOutlookOut.MailAdressIn, parameters.LoginR7751)); //Сюда идентификатор
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
                                }
                                Loggers.Log4NetLogger.Info(new Exception($"Отправка письма {mailLotusOutlookOut.IdMail} на адрес {mail}"));
                            }
                            catch (Exception ex)
                            {
                                Loggers.Log4NetLogger.Error(ex);
                                mailLotusOutlookOut.ErrorMail = $"Письмо не отправлено адресату возникли ошибки во время отправки";
                            }
                        }
                        Mail.SendMailAutoOutput(user, $"Отправка писем произведена!!!", $"Адреса участники рассылки \r\n {string.Join("\r\n", arrayMail)}");
                    }
                    else
                    {
                        Mail.SendMailAutoOutput(user, $"Отправка писем на адрес(а) не возможна {mailLotusOutlookOut.MailAdressOut} !!!", $"Ошибка почты {mailLotusOutlookOut.MailAdressOut} !!!");
                        mailLotusOutlookOut.ErrorMail = $"Почтовый ящик(и) не прошел(и) проверку отправка не возможна!";
                        Loggers.Log4NetLogger.Error(new Exception($"Отправка письма {mailLotusOutlookOut.IdMail} не прошла проверку внешних адресов"));
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
        /// Поиск всех Email Адресов
        /// </summary>
        /// <param name="emailAddress">Тема описание из Lotus</param>
        /// <returns></returns>
        private string[] MailArraySubject(string emailAddress)
        {
            return Regex.Matches(emailAddress, @"([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)").Cast<Match>().Select(m => m.Value).ToArray();
        }
   }
}
