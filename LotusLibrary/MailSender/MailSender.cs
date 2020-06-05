using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Domino;
using EfDatabase.Inventory.Base;
using LotusLibrary.DbConnected;
using LotusLibrary.DxlLotus;
using LotusLibrary.DxlLotus.DocumentGeneration;


namespace LotusLibrary.MailSender
{
   public class MailSender : IDisposable
    {
        /// <summary>
        /// View
        /// </summary>
        public NotesView NotesView { get; set; }
        /// <summary>
        /// Document
        /// </summary>
        public NotesDocument Document { get; set; }
        /// <summary>
        /// Поток Lotus
        /// </summary>
        public NotesStream NotesStream { get; set; }
        private LotusConfig Config { get; set; }
        private LotusConnectedDataBase Db { get; set; }

        public MailSender()
        {
            
            Dispose();
            Config = new LotusConfig();
            Db = new LotusConnectedDataBase(Config.LotusIdFilePassword);

        }

        /// <summary>
        /// Поиск пользователя по табельному номеру в справочниках Lotus
        /// </summary>
        /// <param name="userSql">Идентификатор из БД Инвентаризация</param>
        /// <param name="infoMail">Информация для Log</param>
        public string[] FindUserLotusMail(EfDatabaseXsdLotusUser.UserLotus userSql,string infoMail)
        {
          try
          { 
              var book = Db.LotusConnectedDataBaseServer(Config.LotusServer, "names.nsf");
              var mailUsers = new string[userSql.User.Length];
              var i = 0;
              foreach (var user in userSql.User)
              {
                 var arrayFio = user.Name.Split(' ');
                 var id = user.TabelNumber.Substring(0, 4);
                 var documentAllMail = book.Search(String.Format("Select(Person_LastNameRUS=\"" + arrayFio[0] + "\"&Person_FirstNameRUS=\"" + arrayFio[1] + "\"&Person_MiddleNameRUS=\"" + arrayFio[2] + "\"&@Contains(FullName;\"" + id + "\"))"), null, 0);
                 if (documentAllMail.Count == 0)
                 {
                     documentAllMail = book.Search(String.Format("@Select(EmployeeID=\"" + user.TabelNumber + "\")"), null, 0);
                 }
                 var docMail = documentAllMail.GetFirstDocument();
                 if (docMail == null)
                 {
                     Loggers.Log4NetLogger.Info(new Exception($"Пользователь с табельным номером {user.TabelNumber} не найден в БД Lotus Mosnalog's Address Book {infoMail}"));
                 }
                 while (docMail != null)
                 {
                     string fullName = docMail.GetItemValue(@"FullName")[0];
                     string mailDomain = docMail.GetItemValue(@"MailDomain")[0];
                     mailUsers[i] = String.Concat(fullName.Replace("CN=", "")
                         .Replace("OU=", "")
                         .Replace("O=", ""), "@", mailDomain);
                     docMail = documentAllMail.GetNextDocument(docMail);
                     i++;
                 }
              }
              return mailUsers;
          }
          catch (Exception e)
          {
              Loggers.Log4NetLogger.Error(e);
          }
          return null;
        }

        /// <summary>
        /// Отправка письма пользователям по группе
        /// </summary>
        /// <param name="mailOutlook">Письма заступившие</param>
        /// <param name="arrayUsers">Рассылка пользователям</param>
        public void SendMailIn(MailLotusOutlookIn mailOutlook,string[] arrayUsers)
        {
            Db.LotusConnectedDataBaseServer(Config.LotusServer, Config.LotusMailSend);
            DocumentGenerationAllDxl document = new DocumentGenerationAllDxl(Db.Db);
            document.DocumentGenerationMailMemo(arrayUsers, "От кого: " + mailOutlook.MailAdress + " Тема: " + mailOutlook.SubjectMail, Db.Session.UserName, mailOutlook.Body, mailOutlook.FullPathFile);
            DonloadOnCreateDxlFile download = new DonloadOnCreateDxlFile();
            download.DxlFileSave(Config.PathGenerateScheme, document.Document, typeof(note));
            var noteId = download.ImportDxlFile(Config.PathGenerateScheme, Db.Db);
            if (noteId != null)
            {
                var docSave = Db.Db.GetDocumentByID(noteId);
                docSave.Send(false);
            }
        }
        /// <summary>
        /// Рассылка писем если пришло Html В случае если зашло Html
        /// </summary>
        /// <param name="mailOutlook">Письма заступившие</param>
        /// <param name="arrayUsers">Рассылка пользователям</param>
        public void SendMailMimeHtml(MailLotusOutlookIn mailOutlook, string[] arrayUsers)
        {
            Db.LotusConnectedDataBaseServer(Config.LotusServer, Config.LotusMailSend);
            NotesStream = Db.Db.Parent.CreateStream();
            Document = Db.Db.CreateDocument();
            try
            {
                if (Db.Db == null)
                    throw new InvalidOperationException("Фатальная ошибка нет соединения с сервером!");
                Document.AppendItemValue("Subject", "От кого: " + mailOutlook.MailAdress + " Тема: " + mailOutlook.SubjectMail);
                Document.AppendItemValue("Recipients", Db.Session.UserName);
                Document.AppendItemValue("OriginalTo", Db.Session.UserName);
                Document.AppendItemValue("From", Db.Session.UserName);
                Document.AppendItemValue("OriginalFrom", Db.Session.UserName);
                Document.AppendItemValue("SendTo", arrayUsers);

                var mimeEntity = Document.CreateMIMEEntity();

                var mimeHeader = mimeEntity.CreateHeader("MIME-Version");
                mimeHeader.SetHeaderVal("1.0");

                mimeHeader = mimeEntity.CreateHeader("Content-Type");
                mimeHeader.SetHeaderValAndParams("multipart/mixed");

                var mimeChild = mimeEntity.CreateChildEntity();
                NotesStream.WriteText(mailOutlook.Body);

                mimeChild.SetContentFromText(NotesStream, "text/html;charset=\"utf-8\"", MIME_ENCODING.ENC_NONE);
                mimeChild = mimeEntity.CreateChildEntity();

                if (File.Exists(mailOutlook.FullPathFile))
                {
                    NotesStream.Truncate();
                    NotesStream.Open(mailOutlook.FullPathFile);
                    mimeHeader = mimeChild.CreateHeader("Content-Disposition");
                    mimeHeader.SetHeaderVal($"attachment; filename=\"{mailOutlook.NameFile}\"");
                    mimeChild.SetContentFromBytes(NotesStream, $"application/zip; name=\"{mailOutlook.NameFile}\"", MIME_ENCODING.ENC_IDENTITY_BINARY);
                }
                Document.CloseMIMEEntities(true);
                Db.Db.Parent.ConvertMime = true;
                if (Document.ComputeWithForm(true, false))
                {
                    Document.Save(true, true);
                    Document.Send(false);
                    NotesStream.Truncate();
                    NotesStream.Close();
                }
            }
            catch (Exception e)
            {
                Loggers.Log4NetLogger.Error(new Exception("В рассылке Html письма возникла ошибка!"));
                Loggers.Log4NetLogger.Error(e);
            }
            finally
            {
                if (Document != null)
                    Marshal.ReleaseComObject(Document);
                Document = null;
                if (NotesStream != null)
                    Marshal.ReleaseComObject(NotesStream);
                NotesStream = null;
            }
        }

        /// <summary>
        /// Генерация производного документа ответа БД должна быть открыта
        /// </summary>
        /// <param name="arrayUsers">Кому отправлять</param>
        /// <param name="subject">Тема</param>
        /// <param name="body">Тело документа</param>
        /// <param name="fileFullPathName">Имя файла</param>
        public void SendMailAutoOutput(string[] arrayUsers, string subject, string body, string fileFullPathName = null)
        {
            if (Db.Db != null)
            {
                DocumentGenerationAllDxl document = new DocumentGenerationAllDxl(Db.Db);
                document.DocumentGenerationMailMemo(arrayUsers, subject, Db.Session.UserName, body, fileFullPathName);
                DonloadOnCreateDxlFile download = new DonloadOnCreateDxlFile();
                download.DxlFileSave(Config.PathGenerateScheme, document.Document, typeof(note));
                var noteId = download.ImportDxlFile(Config.PathGenerateScheme, Db.Db);
                if (noteId != null)
                {
                    var docSave = Db.Db.GetDocumentByID(noteId);
                    docSave.Send(false);
                }
            }
            else
            {
                Loggers.Log4NetLogger.Error(new Exception($"База не открыта или пуста для обратной связи ошибки"));
            }
        }
        /// <summary>
        /// Формирование модели для отправки по SMTP протоколу
        /// </summary>
        /// <param name="pathSaveFile">Путь сохранения вложенных в письма файлов на отправку</param>
        public List<MailLotusOutlookOut> SendMailOut(string pathSaveFile)
        {
            var mailLotusOutlookOut = new List<MailLotusOutlookOut>();
            var publicFunction = new PublicFunctionInfo.PublicFunctionInfo();
            Db.LotusConnectedDataBaseServer(Config.LotusServer, Config.LotusMailSend);
            if (Db.Db == null)
                throw new InvalidOperationException("Фатальная ошибка нет соединения с сервером!");
            var docList = new List<string>();
            try
            {
                NotesView = Db.GetViewLotus("$Inbox");
                Document = NotesView.GetFirstDocument();
                while (Document != null)
                {
                    var mail = new MailLotusOutlookOut
                    {
                        IdMail = Document.UniversalID,
                        MailAdressOut = Document.GetItemValue("Subject")[0].ToString(),
                        Body = Document.GetItemValue("Body")[0].ToString()
                    };
                    NotesRichTextItem bodyFileExtractMail = Document.GetFirstItem("Body") as NotesRichTextItem;
                    var nameFileList = publicFunction.ExtractFile(bodyFileExtractMail, pathSaveFile);
                    mail.FullPathListFile = nameFileList != null && nameFileList.Count > 0 ? string.Join(";", nameFileList) : null;
                    mail.MailAdressIn = Document.GetItemValue("From")[0].ToString();
                    docList.Add(Document.UniversalID);
                    mailLotusOutlookOut.Add(mail);
                    Document = NotesView.GetNextDocument(Document);
                }
            }
            catch (Exception ex)
            {
                Loggers.Log4NetLogger.Error(ex);
                throw;
            }
            finally
            {
                if (Document != null)
                    Marshal.ReleaseComObject(Document);
                Document = null;
                if (NotesView != null)
                    Marshal.ReleaseComObject(NotesView);
                NotesView = null;
            }
            foreach (var unId in docList)
            {
                Document = Db.Db.GetDocumentByUNID(unId);
                if (Document == null)
                    throw new InvalidOperationException($"No document {unId}");
                Document.Remove(true);
                Marshal.ReleaseComObject(Document);
            }
            return mailLotusOutlookOut;
        }

        private void ReleaseUnmanagedResources()
        {

        }

        private void Dispose(bool disposing)
        {
            ReleaseUnmanagedResources();
            if (disposing)
            {
                Db?.Dispose();
            }
        }

        public void Dispose()
        {
            Dispose(true);
          //  GC.SuppressFinalize(this);
        }

        ~MailSender()
        {
            Dispose(false);
        }
    }
}
