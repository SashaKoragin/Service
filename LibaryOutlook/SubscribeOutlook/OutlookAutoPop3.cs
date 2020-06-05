using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EfDatabase.Inventory.Base;
using EfDatabase.Inventory.BaseLogic.Select;
using EfDatabase.Inventory.MailLogicLotus;
using EfDatabaseXsdLotusUser;
using LotusLibrary.MailSender;
using MailKit.Net.Pop3;
using MimeKit;


namespace LibraryOutlook.SubscribeOutlook
{
   public class OutlookAutoPop3
   {

        /// <summary>
        /// Пересылка с почты oit на пользователей Lotus
        /// </summary>
        /// <param name="parameters"></param>
        public void StartMessageOit(ConfigFile.ConfigFile parameters)
       {
          try
          {
              int count = 0;
              ZipAttachments zip = new ZipAttachments();
                using (Pop3Client client = new Pop3Client())
                {
                    client.CheckCertificateRevocation = false;
                    client.Connect(parameters.Pop3Address, 995, true);
                    MailSender mail = new MailSender();
                    Loggers.Log4NetLogger.Info(new Exception($"Соединение с сервером eups.tax.nalog.ru установлено (OIT)"));
                    client.Authenticate(parameters.LoginOit, parameters.PasswordOit);
                    Loggers.Log4NetLogger.Info(new Exception($"Пользователь проверен (OIT)"));
                    if (client.IsConnected)
                    {
                        MailLogicLotus mailSave = new MailLogicLotus();
                        SelectSql select = new SelectSql();
                        UserLotus userSqlDefault = select.FindUserGroup(7);
                        int messageCount = client.GetMessageCount();
                        
                        for (int i = 0; i < messageCount; i++)
                        {
                            MimeMessage message = client.GetMessage(i);
                            var messageAttaches = message.Attachments as List<MimeEntity> ?? new List<MimeEntity>();
                            var messageBodyAttaches = new List<MimeEntity>();
                            messageBodyAttaches.AddRange(message.BodyParts);
                            var file = messageBodyAttaches.Where(x =>
                                x.ContentType.MediaType == "image" ||
                                x.ContentType.MediaType == "application").ToList();
                            if (file.Count>0)
                                messageAttaches.AddRange(file);
                            string body;
                            var isHtmlMime = false;
                            if (string.IsNullOrWhiteSpace(message.TextBody))
                            {
                                body = message.HtmlBody;
                                isHtmlMime = true;
                            }
                            else
                            {
                                body = message.TextBody;
                            }
                            var date = message.Date;
                            if (date.Date >= DateTime.Now.Date)
                            {
                                if (!mailSave.IsExistsBdMail(message.MessageId))
                                {
                                    var address = (MailboxAddress)message.From[0];
                                    var mailSender = address.Address;
                                    var nameFile = date.ToString("dd.MM.yyyy_HH.mm.ss") + "_" + mailSender.Split('@')[0] + ".zip";
                                    var fullPath = Path.Combine(parameters.PathSaveArchive, nameFile);
                                    MailLotusOutlookIn mailMessage = new MailLotusOutlookIn()
                                    {
                                        IdMail = message.MessageId,
                                        MailAdressSend = parameters.LoginOit,
                                        SubjectMail = message.Subject,
                                        Body = body,
                                        MailAdress = mailSender,
                                        DateInputServer = date.Date,
                                        NameFile = nameFile,
                                        FullPathFile = fullPath,
                                        FileMail = zip.StartZipArchive(messageAttaches, fullPath)
                                    };
                                    mailSave.AddModelMailIn(mailMessage);
                                    var mailUsers = mail.FindUserLotusMail(select.FindUserOnUserGroup(userSqlDefault, mailMessage.SubjectMail), "(OIT)");
                                    if (isHtmlMime)
                                    {
                                        mail.SendMailMimeHtml(mailMessage, mailUsers);
                                    }
                                    else
                                    {
                                        mail.SendMailIn(mailMessage, mailUsers);
                                    }
                                    count++;
                                    Loggers.Log4NetLogger.Info(new Exception($"УН: {mailMessage.IdMail} Дата/Время: {date} От кого: {mailMessage.MailAdress}"));
                                }
                            }
                            else
                            {
                              //Удаление сообщения/письма
                              client.DeleteMessage(i);
                            }
                        }
                        Loggers.Log4NetLogger.Info(new Exception("Количество пришедшей (OIT) почты:" + count));
                    }
                    mail.Dispose();
                    client.Disconnect(true);
                }
                //Очистить временную папку с файлами
                zip.DropAllFileToPath(parameters.PathSaveArchive);
                foreach (FileInfo file in new DirectoryInfo(parameters.PathSaveArchive).GetFiles())
                {
                    Loggers.Log4NetLogger.Info(new Exception($"Наименование удаленных файлов: {file.FullName}"));
                    file.Delete();
                }
                Loggers.Log4NetLogger.Info(new Exception("Очистили папку от файлов (OIT)!"));
                Loggers.Log4NetLogger.Info(new Exception("Перерыв 5 минут (OIT)"));
          }
          catch (Exception x)
          {
              Loggers.Log4NetLogger.Debug(x);
          }
       }

       /// <summary>
       /// Пересылка с почты r7751 на пользователей Lotus
       /// </summary>
       /// <param name="parameters">Параметры почты</param>
       public void StartMessageR7751(ConfigFile.ConfigFile parameters)
       {
            try
            {
                int count = 0;
                ZipAttachments zip = new ZipAttachments();
                using (Pop3Client client = new Pop3Client())
                {
                    client.CheckCertificateRevocation = false;
                    client.Connect(parameters.Pop3Address, 995, true);
                    MailSender mail = new MailSender();
                    Loggers.Log4NetLogger.Info(new Exception($"Соединение с сервером eups.tax.nalog.ru установлено (R7751)"));
                    client.Authenticate(parameters.LoginR7751, parameters.PasswordR7751);
                    Loggers.Log4NetLogger.Info(new Exception($"Пользователь проверен (R7751)"));
                    if (client.IsConnected)
                    {
                        MailLogicLotus mailSave = new MailLogicLotus();
                        SelectSql select = new SelectSql();
                        int messageCount = client.GetMessageCount();
                        for (int i = 0; i < messageCount; i++)
                        {
                            MimeMessage message = client.GetMessage(i);
                            var messageAttaches = message.Attachments as List<MimeEntity> ?? new List<MimeEntity>();
                            var messageBodyAttaches = new List<MimeEntity>();
                            messageBodyAttaches.AddRange(message.BodyParts);
                            var file = messageBodyAttaches.Where(x =>
                                x.ContentType.MediaType == "image" ||
                                x.ContentType.MediaType == "application").ToList();
                            if (file.Count > 0)
                                messageAttaches.AddRange(file);
                            var date = message.Date;
                            if (date.Date >= DateTime.Now.Date)
                            {
                                if (!mailSave.IsExistsBdMail(message.MessageId))
                                {
                                    var address = (MailboxAddress)message.From[0];
                                    var mailSender = address.Address;
                                    string body;
                                    var isHtmlMime = false;
                                    if (string.IsNullOrWhiteSpace(message.TextBody))
                                    {
                                        body = message.HtmlBody;
                                        isHtmlMime = true;
                                    }
                                    else
                                    {
                                        body = message.TextBody;
                                    }
                                    var nameFile = date.ToString("dd.MM.yyyy_HH.mm.ss") + "_" + mailSender.Split('@')[0] + ".zip";
                                        var fullPath = Path.Combine(parameters.PathSaveArchive, nameFile);
                                        MailLotusOutlookIn mailMessage = new MailLotusOutlookIn()
                                        {
                                            IdMail = message.MessageId,
                                            MailAdressSend = parameters.LoginR7751,
                                            SubjectMail = message.Subject,
                                            Body = body,
                                            MailAdress = mailSender,
                                            DateInputServer = date.Date,
                                            NameFile = nameFile,
                                            FullPathFile = fullPath,
                                            FileMail = zip.StartZipArchive(messageAttaches, fullPath)
                                        };
                                        mailSave.AddModelMailIn(mailMessage);
                                        //Конец блока сохранения письма
                                        var userSql = select.FindUserOnUserGroup(null, mailMessage.SubjectMail);
                                        //Если нашли пользователя в БД
                                        if (userSql?.User != null && userSql.User.Length > 0)
                                        {
                                            var mailUsers = mail.FindUserLotusMail(userSql, "(R7751)");
                                            if (mailUsers.Length > 0)
                                            {
                                                if (isHtmlMime)
                                                {
                                                    mail.SendMailMimeHtml(mailMessage, mailUsers);
                                                }
                                                else
                                                {
                                                    mail.SendMailIn(mailMessage, mailUsers);
                                                }
                                            }
                                            count++;
                                            Loggers.Log4NetLogger.Info(new Exception($"УН: {mailMessage.IdMail} Дата/Время: {date} От кого: {mailMessage.MailAdress}"));
                                        }
                                }
                            }
                            else
                            {
                                //Удаление сообщения/письма
                                client.DeleteMessage(i);
                            }
                        }
                        Loggers.Log4NetLogger.Info(new Exception("Количество пришедшей почты (R7751):" + count));
                    }
                    mail.Dispose();
                    client.Disconnect(true);
                }
                //Очистить временную папку с файлами
                zip.DropAllFileToPath(parameters.PathSaveArchive);
                Loggers.Log4NetLogger.Info(new Exception("Очистили папку от файлов! (R7751)"));
                Loggers.Log4NetLogger.Info(new Exception("Перерыв 5 минут (R7751)"));
            }
            catch (Exception x)
            {
                Loggers.Log4NetLogger.Debug(x);
            }
       }
   }
}
