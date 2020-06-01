using System.IO;
using Domino;
using LibraryOutlook.SubscribeOutlook;
using LotusLibrary.DbConnected;
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
            outlook.StartMessageOit(new LibraryOutlook.ConfigFile.ConfigFile());
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
            var docs = lotusMail.Search("Select($MessageID= \"<OFBE849B44.93AAB6F7-ON43258515.0034B243-43258515.0034BD93@LocalDomain>\")", null, 0);
            NotesDXLExporter export = db.Session.CreateDXLExporter();
            var t = export.Export(docs.GetFirstDocument());
            File.WriteAllText(filename, t);
        }

        [TestMethod()]
        public void StartMessageTest1()
        {
            LibraryOutlook.ConfigFile.ConfigFile Parameters = new LibraryOutlook.ConfigFile.ConfigFile();
            OutlookAutoPop3 t = new OutlookAutoPop3();
            t.StartMessageOit(Parameters);
        }
        [TestMethod()]
        public void FindUserLotus()
        {
            var smtp = new OutlookAutoSmtp();
            smtp.SendSmtpMessage(new LibraryOutlook.ConfigFile.ConfigFile());
        }
    }
}