using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Domino;


namespace LotusLibrary.DxlLotus.DocumentGeneration
{
   public class DocumentGenerationAllDxl:IDisposable
    {

        public note Document { get; set; }

        public DocumentGenerationAllDxl(NotesDatabase database)
        {
            Document = new note();
            Document.@class = noteClass.document;
            Document.version = "6.5";
            Document.maintenanceversion = "5.0";
            Document.replicaid = database.ReplicaID;
            Document.updatedby = null;
            Document.revisions = null;
            Document.wassignedby = null;
            Document.textproperties = null;
            Document.noteinfo.noteid = "0";
            Document.noteinfo.sequence = "0";
            Document.noteinfo.created.datetime.Text = new List<string>(new[] { DateTime.Now.ToString(CultureInfo.InvariantCulture) });
            Document.noteinfo.modified.datetime.Text = new List<string>(new string[] {  });
            Document.noteinfo.revised.datetime.Text = new List<string>(new[] { DateTime.Now.ToString(CultureInfo.InvariantCulture) });
            Document.noteinfo.lastaccessed.datetime.Text = new List<string>(new[] { DateTime.Now.ToString(CultureInfo.InvariantCulture) });
            Document.noteinfo.addedtofile.datetime.Text = new List<string>(new string[] {  });

        }
        /// <summary>
        /// Генерация документа Почты
        /// </summary>
        /// <param name="sendTo">Кому отправлять</param>
        /// <param name="subject">Тема</param>
        /// <param name="sender">Отправитель</param>
        /// <param name="body">Тело документа</param>
        /// <param name="fileName">Имя файла</param>
        public void DocumentGenerationMailMemo(string[] sendTo, string subject, string sender, string body, string fileName)
        {
            Document.Items.Add(new item() { name = "Subject", Item = GenerateText(new[] { subject }) });
            Document.Items.Add(new item() { name = "Recipients", Item = GenerateText(new[] { sender }) });
            Document.Items.Add(new item() { name = "PostedDate", Item = GenerateText(new[] { DateTime.Now.ToString(CultureInfo.InvariantCulture)}) });
            Document.Items.Add(new item() { name = "Form", Item = GenerateText(new[] { "Memo"}) });
            Document.Items.Add(new item() { name = "AltFrom", Item = new textlist() { text = new List<text>(new[] { GenerateText(new string[] { }) }) } });
            Document.Items.Add(new item() { name = "Logo", Item = GenerateText(new[] { "StdNotesLtrGateway"}) });
            Document.Items.Add(new item() { name = "useApplet", Item = GenerateText(new[] { "True"}) });
            Document.Items.Add(new item() { name = "DefaultMailSaveOptions", Item = GenerateText(new[] { "1"}) });
            Document.Items.Add(new item() { name = "Query_String", Item = new textlist() { text = new List<text>(new[] { GenerateText(new string[] { }) }) } });
            Document.Items.Add(new item() { name = "Encrypt", Item = GenerateText(new[] { "0"}) });
            Document.Items.Add(new item() { name = "Sign", Item = GenerateText(new[] { "0"}) });
            Document.Items.Add(new item() { name = "BlindCopyTo", names = itemNames.@true, Item = GenerateText(new string[] { }) });
            Document.Items.Add(new item() { name = "WebSubject", names = itemNames.@true, Item = GenerateText(new[] { sender }) });
            Document.Items.Add(new item() { name = "FaxToList", names = itemNames.@true, Item = GenerateText(new string[] { }) });
            Document.Items.Add(new item() { name = "EnterSendTo", Item = new textlist() { text = new List<text>(new[] { GenerateText(new[] { sender }) }) } });
            Document.Items.Add(new item() { name = "EnterCopyTo", Item = new textlist() { text = new List<text>(new[] { GenerateText(new string[] { }) }) } });
            Document.Items.Add(new item() { name = "EnterBlindCopyTo", Item = new textlist() { text = new List<text>(new[] { GenerateText(new string[] { }) }) } });
            Document.Items.Add(new item() { name = "OriginalTo", Item = GenerateText(new[] { sender }) });
            Document.Items.Add(new item() { name = "SendTo", Item = GenerateTextList(sendTo) });
            Document.Items.Add(new item() { name = "OriginalFrom", Item = GenerateText(new[] { sender }) });
            Document.Items.Add(new item() { name = "From", Item = GenerateText(new []{ sender }) });
            Document.Items.Add(new item()
                {
                    name = "Body",
                    sign = itemSign.@true,
                    seal = itemSeal.@true,
                    Item =
                        new richtext()
                        {
                            Items = new List<object>()
                            {
                                GenerateParagraphNewCollection(string.IsNullOrWhiteSpace(body) ? "Отсутствует описание!!!" : body),
                                GenerateParagraphPicture(fileName)
                            }
                        }
                });
                if (File.Exists(fileName))
            {
                Document.Items.Add(new item()
                {
                    name = "$FILE",
                    summary = itemSummary.@true,
                    sign = itemSign.@true,
                    seal = itemSeal.@true,
                    Item = SetFileDxlFile(fileName)
                });
            }
        }

        /// <summary>
        /// Генерация текстового поля
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private text GenerateText(string[] text)
        {
           return new text() {Text = new List<string>(text)};
        }
        /// <summary>
        /// Генерация параграфа с иконкой или пустого 
        /// </summary>
        /// <param name="filePathName">Путь к файлу</param>
        /// <returns></returns>
        private par GenerateParagraphPicture(string filePathName)
        {
            if (File.Exists(filePathName))
            {
                return new par()
                {
                    Items = new List<object>()
                    {
                        new @break(),
                        GetPicture(filePathName)
                    }
                };
            }
            return new par();
        }

        /// <summary>
        /// Извлекаем файл в Base64
        /// </summary>
        /// <param name="filePathName">Путь к файлу</param>
        /// <returns></returns>
        private @object SetFileDxlFile(string filePathName)
        {
            using (FileStream output = new FileStream(filePathName, FileMode.Open))
            {
                int length = Convert.ToInt32(output.Length);
                byte[] data = new byte[length];
                output.Read(data, 0, length);
                return new @object()
                {
                    file = new file()
                    {
                        hosttype = fileHosttype.msdos,
                        name = Path.GetFileName(filePathName),
                        filedata = new filedata() {Text = new List<string>(new[] {Convert.ToBase64String(data)})}
                    }
                };
            }
        }

        /// <summary>
        /// Генерация параграфа с разделителем \r\n
        /// </summary>
        /// <param name="text">Текст с разделителями \r\n</param>
        /// <returns></returns>
        private par GenerateParagraphNewCollection(string text)
        {
            string[] stringSeparators = new string[] { "\r\n" };
            var strParagraph = text.Split(stringSeparators, StringSplitOptions.None);
            var paragraph = new par() {Items = new List<object>()};
            foreach (var paragraphText in strParagraph)
            {
                paragraph.Items.Add(new run()
                {
                    Items = new List<object>() {new font() {size = "14pt"}},
                    Text = new List<string>(new[] { paragraphText })
                });
                paragraph.Items.Add(new @break());
            }
            return paragraph;
        }

        private textlist GenerateTextList(string[] arrayText)
        {
            var i = 0;
            var textArray = new text[arrayText.Length];
            var textList = new textlist() {text = new List<text>() };
            foreach (var text in arrayText)
            {
                textArray[i] = new text();
                textArray[i].Text = new List<string>(new []{text});
                textList.text.Add( textArray[i]);
                i++;
            }
            return textList;
        }


        /// <summary>
        /// Рисуем иконку файла
        /// </summary>
        /// <param name="filePathName">Путь к файлу</param>
        /// <returns></returns>
        private attachmentref GetPicture(string filePathName)
        {
            var nameFile = Path.GetFileName(filePathName);
                using (var font = new Font("Tahoma", 8))
                {
                    SizeF textSize = new SizeF();
                    if (filePathName != null)
                        using (var iconBitmap = Icon.ExtractAssociatedIcon(filePathName))
                        {
                            float yShift = 0f;
                            if (iconBitmap != null)
                            {
                                using (var tempBitmap = new Bitmap(iconBitmap.Width, iconBitmap.Height))
                                {
                                    using (var drawing = Graphics.FromImage(tempBitmap))
                                    {
                                        var sz = drawing.MeasureString(filePathName, font);
                                        textSize.Width = iconBitmap.Width + sz.Width + 10;
                                        textSize.Height = Math.Max(sz.Height, iconBitmap.Height);
                                        if (sz.Height < iconBitmap.Height)
                                            yShift = (iconBitmap.Height - sz.Height) / 2;
                                    }
                                }

                                using (var img = new Bitmap((int) textSize.Width, (int) textSize.Height))
                                {
                                    using (var drawing = Graphics.FromImage(img))
                                    {
                                        drawing.Clear(Color.White);
                                        drawing.DrawImage(iconBitmap.ToBitmap(), Point.Empty);
                                        using (var textBrush = new SolidBrush(Color.Black))
                                        {
                                            drawing.DrawString(nameFile, font, textBrush, iconBitmap.Width + 10,
                                                yShift);
                                            drawing.Save();
                                        }
                                    }

                                    using (var ms = new MemoryStream())
                                    {
                                        img.Save(ms, ImageFormat.Gif);
                                        var reattachment = new attachmentref()
                                        {
                                            name = nameFile,
                                            displayname = nameFile,
                                            caption = nameFile,
                                            Items = new List<object>()
                                            {
                                                new picture()
                                                {
                                                    height = $"{textSize.Height}px",
                                                    width = $"{textSize.Width}px",
                                                    Item = new gif()
                                                    {
                                                        Text = new List<string>(
                                                            new[] {Convert.ToBase64String(ms.ToArray())})
                                                    }
                                                }
                                            }
                                        };
                                        return reattachment;
                                    }
                                }
                            }

                           
                        }
                    return null;
            }
        }
        public void Dispose()
        {

        }
    }
}
