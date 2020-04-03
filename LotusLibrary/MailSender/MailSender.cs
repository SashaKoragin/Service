using System;
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
        private LotusConfig Config { get; set; }
        private LotusConnectedDataBase Db { get; set; }

        public MailSender()
        {
            Dispose();
            Db?.Dispose();
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
        public void SendMail(MailOutlook mailOutlook,string[] arrayUsers)
        {
            Db.LotusConnectedDataBaseServer(Config.LotusServer, Config.LotusMailSend);
            DocumentGenerationAllDxl document = new DocumentGenerationAllDxl(Db.Db);
            document.DocumentGenerationMailMemo(arrayUsers, "От кого: " + mailOutlook.MailAdress + " Тема: " + mailOutlook.SubjectMail, Db.Session.UserName, mailOutlook.Body, mailOutlook.FullPathFile);
            DonloadOnCreateDxlFile download = new DonloadOnCreateDxlFile();
            download.DxlFileSave(Config.PathGenerateScheme, document.Document, typeof(note));
            var noteId = download.ImportDxlFile(Config.PathGenerateScheme, Db.Db);
            var docSave = Db.Db.GetDocumentByID(noteId);
            docSave.Send(false);
            Db.Dispose();
        }

        public void Dispose()
        {

            Config = null;
            Db = null;
        }
    }
}
