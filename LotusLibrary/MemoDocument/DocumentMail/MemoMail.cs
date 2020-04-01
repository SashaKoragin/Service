using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Domino;
using LotusLibrary.Mime;

namespace LotusLibrary.MemoDocument.DocumentMail
{
   public class MemoMail
    {

        public NotesDocument Document { get; set; }

        public MemoMail(NotesDocument document)
        {
            Document = document;
        }
        /// <summary>
        /// Создание полей документа без Body 
        /// </summary>
        /// <param name="sendTo">Список рассылки</param>
        /// <param name="subject">Тема</param>
        /// <param name="sender">Отправитель</param>
        public void DocumentAddMemo(string[] sendTo, string subject, string sender)
        {
            Document.ReplaceItemValue("Subject", subject);
            Document.ReplaceItemValue("SendTo", sendTo);
            Document.ReplaceItemValue("Recipients", sender);
            Document.ReplaceItemValue("Principal", sender);
            Document.ReplaceItemValue("From", sender);
            Document.ReplaceItemValue("PostedDate", $"{DateTime.Now}");
        }

        /// <summary>
        /// Добавление файла в Body
        /// </summary>
        /// <param name="notesStream">Поток байт файл</param>
        /// <param name="filename">Имя файла</param>
        /// <param name="mimeFile">MIME файла</param>

        public void AddDocumentFileBody(NotesStream notesStream, string filename,  AllMime.Mime mimeFile)
        {
            var mime = Document.CreateMIMEEntity();
            mime.CreateHeader("Content-Type").SetHeaderValAndParams("multipart/mixed");
            mime.CreateHeader("Content-Language").SetHeaderValAndParams("ru");
            mime.CreateHeader("Content-Disposition").SetHeaderValAndParams("attachment; filename=\"" + filename + "\"");
            mime.SetContentFromBytes(notesStream, mimeFile+"; charset=UTF-8; name=\"" + filename + "\"", MIME_ENCODING.ENC_IDENTITY_BINARY);
        }
        /// <summary>
        /// Добавление текста
        /// </summary>
        /// <param name="textBody">Текст</param>
        /// <param name="styleBody">Стиль</param>
        public void AddBodyText(string textBody = "", NotesRichTextStyle styleBody = null)
        {
            var richText = Document.CreateRichTextItem("Body");
            richText.AppendStyle(styleBody);
            richText.AppendText(textBody);
            richText.AddNewLine();
        }

        /// <summary>
        /// Сохранение и отправка
        /// </summary>
        public void SaveAndSend()
        {
            var isWithForm = Document.ComputeWithForm(true, true);
            if (isWithForm)
            {
                Document.CloseMIMEEntities(true);
                Document.Save(true, true);
                Document.Send(false);
                return;
            }
            Loggers.Log4NetLogger.Error(new Exception("Ошибка при сохранении документа!"));
        }
    }
}
