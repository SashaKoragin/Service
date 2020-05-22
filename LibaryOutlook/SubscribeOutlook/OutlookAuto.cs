using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using EfDatabase.Inventory.Base;
using EfDatabase.Inventory.BaseLogic.Select;
using EfDatabase.Inventory.MailLogicLotus;
using EfDatabaseXsdLotusUser;
using LotusLibrary.MailSender;
using OpenPop.Mime;
using OpenPop.Pop3;
using Message = OpenPop.Mime.Message;


namespace LibraryOutlook.SubscribeOutlook
{
   public class OutlookAuto
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
                using (Pop3Client client = new Pop3Client())
                {
                    client.Connect(parameters.Pop3Address, 995, true);
                    MailSender mail = new MailSender();
                    Loggers.Log4NetLogger.Info(new Exception($"Соединение с сервером eups.tax.nalog.ru установлено (OIT)"));
                    client.Authenticate(parameters.LoginOit, parameters.PasswordOit, AuthenticationMethod.UsernameAndPassword);
                    Loggers.Log4NetLogger.Info(new Exception($"Пользователь проверен (OIT)"));
                    if (client.Connected)
                    {
                        MailLogicLotus mailSave = new MailLogicLotus();
                        SelectSql select = new SelectSql();
                        UserLotus userSqlDefault = select.FindUserGroup(7);
                        int messageCount = client.GetMessageCount();
                        for (int i = messageCount; i > 0; i--)
                        {
                            Message message = client.GetMessage(i);
                            var date = message.Headers.DateSent;
                            if (date.Date >= DateTime.Now.Date)
                            {
                                if (!mailSave.IsExistsBdMail(message.Headers.MessageId))
                                {
                                    MessagePart mpPlain = message.FindFirstPlainTextVersion() ?? message.FindFirstHtmlVersion();
                                    string body = "";
                                    if (mpPlain != null)
                                    {
                                        Encoding enc = mpPlain.BodyEncoding;
                                        body = enc.GetString(mpPlain.Body); //получаем текст сообщения
                                    }
                                    ZipAttachments zip = new ZipAttachments();
                                    var nameFile = date.ToString("dd.MM.yyyy_HH.mm.ss") + "_" + message.Headers.ReturnPath.Address.Split('@')[0] + ".zip";
                                    var fullPath = Path.Combine(parameters.PathSaveArchive, nameFile);
                                    MailOutlook mailMessage = new MailOutlook()
                                    {
                                        IdMail = message.Headers.MessageId,
                                        MailAdressSend = parameters.LoginOit,
                                        SubjectMail = message.Headers.Subject,
                                        Body = body,
                                        MailAdress = message.Headers.ReturnPath.Address,
                                        DateInputServer = date,
                                        NameFile = nameFile,
                                        FullPathFile = fullPath,
                                        FileMail = zip.StartZipArchive(message.FindAllAttachments(),fullPath)
                                    };
                                    mailSave.AddModel(mailMessage);
                                    
                                    if (mailMessage.SubjectMail != null)
                                    {
                                        UserLotus userSql = null;
                                        if (mailMessage.SubjectMail.Length <= 64 && Regex.IsMatch(mailMessage.SubjectMail, @"[А-я]"))
                                        {
                                            //Ищем группу по наименованию если текст
                                            userSql = select.FindUserGroup(null,mailMessage.SubjectMail);
                                        }
                                        if (mailMessage.SubjectMail.Length <= 2 && Regex.IsMatch(mailMessage.SubjectMail, @"\d"))
                                        {
                                            //Ищем группу по номеру группы
                                            userSql = select.FindUserGroup(Convert.ToInt32(mailMessage.SubjectMail));
                                        }
                                        if (mailMessage.SubjectMail.Length > 2 && mailMessage.SubjectMail.Length <= 32 && Regex.IsMatch(mailMessage.SubjectMail, @"\d"))
                                        {
                                            //Ищем пользователя
                                            userSql = select.FindUserSqlKey(mailMessage.SubjectMail);
                                        }
                                        //Если нашли не дефектных пользователей
                                        if (userSql?.User != null && userSql.User.Length>0)
                                        {
                                            userSqlDefault = userSql;
                                        }
                                        else
                                        {
                                            Loggers.Log4NetLogger.Info(new Exception($"Пользователь или группы с идентификатором {mailMessage.SubjectMail} не найден в БД Sql Инвентаризация идет рассылка на стандартную группу (OIT)"));
                                        }
                                    }
                                    var mailUsers = mail.FindUserLotusMail(userSqlDefault, "(OIT)");
                                    mail.SendMail(mailMessage, mailUsers);
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
                    client.Disconnect();
                }
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
                using (Pop3Client client = new Pop3Client())
                {
                    client.Connect(parameters.Pop3Address, 995, true);
                    MailSender mail = new MailSender();
                    Loggers.Log4NetLogger.Info(new Exception($"Соединение с сервером eups.tax.nalog.ru установлено (R7751)"));
                    client.Authenticate(parameters.LoginR7751, parameters.PasswordR7751, AuthenticationMethod.UsernameAndPassword);
                    Loggers.Log4NetLogger.Info(new Exception($"Пользователь проверен (R7751)"));
                    if (client.Connected)
                    {
                        MailLogicLotus mailSave = new MailLogicLotus();
                        SelectSql select = new SelectSql();
                        int messageCount = client.GetMessageCount();
                        for (int i = messageCount; i > 0; i--)
                        {
                            Message message = client.GetMessage(i);
                            var date = message.Headers.DateSent;
                            if (date.Date >= DateTime.Now.Date)
                            {
                                if (!mailSave.IsExistsBdMail(message.Headers.MessageId))
                                {
                                    var messages = message.Headers.Subject;
                                    if (messages != null)
                                    {
                                        //Этот блок к сохранению письма в БД но в данном примере не сохраняем
                                        MessagePart mpPlain = message.FindFirstPlainTextVersion() ?? message.FindFirstHtmlVersion();
                                        string body = "";
                                        if (mpPlain != null)
                                        {
                                            Encoding enc = mpPlain.BodyEncoding;
                                            body = enc.GetString(mpPlain.Body); //получаем текст сообщения
                                        }
                                        ZipAttachments zip = new ZipAttachments();
                                        var nameFile = date.ToString("dd.MM.yyyy_HH.mm.ss") + "_" + message.Headers.ReturnPath.Address.Split('@')[0] + ".zip";
                                        var fullPath = Path.Combine(parameters.PathSaveArchive, nameFile);
                                        MailOutlook mailMessage = new MailOutlook()
                                        {
                                            IdMail = message.Headers.MessageId,
                                            MailAdressSend = parameters.LoginR7751,
                                            SubjectMail = message.Headers.Subject,
                                            Body = body,
                                            MailAdress = message.Headers.ReturnPath.Address,
                                            DateInputServer = date,
                                            NameFile = nameFile,
                                            FullPathFile = fullPath,
                                            FileMail = zip.StartZipArchive(message.FindAllAttachments(), fullPath)
                                        };
                                        mailSave.AddModel(mailMessage);
                                        //Конец блока сохранения письма
                                        UserLotus userSql = null;
                                        if (messages.Length <= 64 && Regex.IsMatch(messages, @"[А-я]"))
                                        {
                                            //Ищем группу по наименованию если текст
                                            userSql = select.FindUserGroup(null, messages);
                                        }
                                        if (messages.Length <= 2 && Regex.IsMatch(messages, @"\d"))
                                        {
                                            //Ищем группу по номеру группы
                                            userSql = select.FindUserGroup(Convert.ToInt32(messages));
                                        }
                                        if (messages.Length > 2 && messages.Length <= 32 && Regex.IsMatch(messages, @"\d"))
                                        {
                                            //Ищем пользователя
                                            userSql = select.FindUserSqlKey(messages);
                                        }
                                        //Если нашли пользователя в БД
                                        if (userSql?.User != null && userSql.User.Length > 0)
                                        {
                                            var mailUsers = mail.FindUserLotusMail(userSql, "(R7751)");
                                            if (mailUsers.Length > 0)
                                            {
                                                mail.SendMail(mailMessage, mailUsers);
                                            }
                                            count++;
                                            Loggers.Log4NetLogger.Info(new Exception($"УН: {mailMessage.IdMail} Дата/Время: {date} От кого: {mailMessage.MailAdress}"));
                                        }
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
                    client.Disconnect();
                }
                foreach (FileInfo file in new DirectoryInfo(parameters.PathSaveArchive).GetFiles())
                {
                    Loggers.Log4NetLogger.Info(new Exception($"Наименование удаленных файлов: {file.FullName}"));
                    file.Delete();
                }
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
