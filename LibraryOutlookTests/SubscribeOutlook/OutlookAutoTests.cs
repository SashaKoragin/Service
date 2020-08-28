using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Domino;
using Domino;
using EfDatabase.Inventory.Base;
using LibraryOutlook.SubscribeOutlook;
using LotusLibrary.DbConnected;
using LotusLibrary.DxlLotus;
using LotusLibrary.DxlLotus.DocumentGeneration;
using LotusLibrary.MailSender;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace LibraryOutlookTests.SubscribeOutlook
{
    [TestClass()]
    public class OutlookAutoTests
    {
        [TestMethod()]
        public void StartMessageTest()
        {
            OutlookAutoPop3 outlook = new OutlookAutoPop3();
            outlook.StartMessageR7751(new LibraryOutlook.ConfigFile.ConfigFile());
        }
        [TestMethod()]
        public void TestLotusServer()
        {
            MailSender mail = new MailSender();
            mail.SendMailIn(null, null);
        }

        [TestMethod()]
        public void TestDocumentView()
        {
            //LotusConnectedDataBase db = new LotusConnectedDataBase(); 
            //var lotusMail = db.LotusConnectedDataBaseServer("12345", "Lotus7751/I7751/R77/МНС", "mail\\акоряг.nsf");

            //DocumentGenerationAllDxl document = new DocumentGenerationAllDxl(lotusMail);
            //document.DocumentGenerationMailMemo("D:\\Скан согласования.pdf");
            //DonloadOnCreateDxlFile download = new DonloadOnCreateDxlFile();
            //download.DxlFileSave("D:\\file.dxl", document.Document, typeof(note));
            //var noteId = download.ImportDxlFile("D:\\file.dxl", lotusMail);
            //var docSave = lotusMail.GetDocumentByID(noteId);
            //docSave.Send(false);
        }

        [TestMethod()]
        public void Dxl()
        {
            LotusConnectedDataBase db = new LotusConnectedDataBase("12345");
            var lotusMail = db.LotusConnectedDataBaseServer("Lotus7751/I7751/R77/МНС", "mail\\акоряг.nsf");
            var streame = db.Session.CreateStream();
            var filename = $"d:\\Test2.dxl";
            var docs = lotusMail.Search("Select($MessageID= \"<OF48F1D289.7607B3B0-ON4325857E.0041458C-4325857E.004145A7@LocalDomain>\")", null, 0);
            NotesDXLExporter export = db.Session.CreateDXLExporter();
            var t = export.Export(docs.GetFirstDocument());
            File.WriteAllText(filename, t);
        }

        [TestMethod()]
        public void StartMessageTest1()
        {
            LibraryOutlook.ConfigFile.ConfigFile Parameters = new LibraryOutlook.ConfigFile.ConfigFile();
            OutlookAutoPop3 t = new OutlookAutoPop3();
            t.StartMessageR7751(Parameters);
        }
        [TestMethod()]
        public void FindUserLotus()
        {
            //var str = "ваываываCN=Михаил Витальевич Мочалов/OU=I7751а/OU=R77/O=МНвСвавыавыаыв";

            //var math = Regex.Match(str, @"CN=(.+)МНС");
            //var fio = math.Value;
            //var fioArray = fio.Split(' ');
            //var fioSelectDb = $"{fioArray[1]} {fioArray[2]} {fioArray[0]}";
             var smtp = new OutlookAutoSmtp();
              smtp.SendSmtpMessage(new LibraryOutlook.ConfigFile.ConfigFile());
        }

        [TestMethod()]
        public void TestAddDocumentReplace()
        {
            var session = new NotesSession();
            session.Initialize("12345");
            var HtmlBody = session.CreateStream();
            var path = @"D:\Вложение 10.zip";
            var mailBody = @"<div><br/>
                               </div><div>
                    <br/>
                    </div>
                    <div class=""c611af8754f78aa8normalize"">
                    <div><div><div><div>Добрый день!</div>
                    <div>Реквизиты:</div>
                    <div>Валюта получаемого перевода: Рубли(RUB)</div>
                    <div>Получатель: НОВИКОВ АЛЕКСЕЙ ВАЛЕРЬЕВИЧ</div>
                    <div>Номер счёта: 
                    <span class=""1f1ea193f6735cf0wmi-callto"">40817810038180282453</span>
                    </div><div>Банк получателя: ПАО СБЕРБАНК</div>
                    <div>БИК: 
                    <span class=""1f1ea193f6735cf0wmi-callto"">044525225</span>
                    </div><div>Корр.счёт: <span class=""1f1ea193f6735cf0wmi-callto"">30101810400000000225</span></div>
                    <div>ИНН: <span class=""1f1ea193f6735cf0wmi-callto"">7707083893</span>
                    </div><div>КПП: <span class=""1f1ea193f6735cf0wmi-callto"">773643001</span>
                    </div><div>SWIFT-код: SABRRUMM</div><div> </div><div>С уважением,</div>
                    <div>Алексей Новиков</div></div></div></div></div><div><br /></div><div>
                    <br /></div><div><br /></div><div><br /></div>
                    <div class=""86e3a5cbc83439f6js-compose-signature""></div><div><br />
                    </div>";
            var db = session.GetDatabase("Lotus7751/I7751/R77/МНС", "mail\\акоряг.nsf", false);
            var document = db.CreateDocument();
            //  document.ComputeWithForm(true,)
            document.AppendItemValue("Subject", "Тема");
            document.AppendItemValue("Recipients", "Алекcандр Сергеевич Корягин/I7751/R77/МНС@MNS_R77");
            document.AppendItemValue("OriginalTo", "Алекcандр Сергеевич Корягин/I7751/R77/МНС@MNS_R77");
            document.AppendItemValue("From", "Алекcандр Сергеевич Корягин/I7751/R77/МНС@MNS_R77");
            document.AppendItemValue("OriginalFrom", "Алекcандр Сергеевич Корягин/I7751/R77/МНС@MNS_R77");
            document.AppendItemValue("SendTo", "Алекcандр Сергеевич Корягин/I7751/R77/МНС@MNS_R77");


            //  document.AppendItemValue("Body","");
            var MIMEEntity = document.CreateMIMEEntity("Body");

            var MimeHeader = MIMEEntity.CreateHeader("MIME-Version");
            MimeHeader.SetHeaderVal("1.0");

            MimeHeader = MIMEEntity.CreateHeader("Content-Type");
            MimeHeader.SetHeaderValAndParams("multipart/mixed");

            var MimeChilds = MIMEEntity.CreateChildEntity();
            HtmlBody.WriteText(mailBody);
            
            MimeChilds.SetContentFromText(HtmlBody, "text/html;charset=\"utf-8\"", Domino.MIME_ENCODING.ENC_NONE);
            MimeChilds = MIMEEntity.CreateChildEntity();


            
            HtmlBody = session.CreateStream();
            HtmlBody.Open(path, "");
            MimeHeader = MimeChilds.CreateHeader("Content-Disposition");

            MimeHeader.SetHeaderVal("attachment; filename=\"16.03.2020_09.35.31_e.safonov.r7700.zip\"");
            MimeChilds.SetContentFromBytes(HtmlBody, "application/zip; name=\"16.03.2020_09.35.31_e.safonov.r7700.zip\"", Domino.MIME_ENCODING.ENC_IDENTITY_BINARY);
           
          
            document.CloseMIMEEntities(true);
            
            session.ConvertMime = true;
            document.CloseMIMEEntities(true);
            if (document.ComputeWithForm(true, false))
            {
                document.Save(true, true);
                document.Send(false);
            }
        }

    }
}