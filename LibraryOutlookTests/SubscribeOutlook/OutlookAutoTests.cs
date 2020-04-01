using LibraryOutlook.SubscribeOutlook;
using System.IO;
using System.Text.RegularExpressions;
using Domino;
using LotusLibrary.DbConnected;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using LotusLibrary.MailSender;
using EfDatabase.Inventory.BaseLogic.Select;


namespace LibraryOutlook.SubscribeOutlook.Tests
{
    [TestClass()]
    public class OutlookAutoTests
    {
        [TestMethod()]
        public void StartMessageTest()
        {
            OutlookAuto outlook = new OutlookAuto();
            outlook.StartMessageOit(new ConfigFile.ConfigFile());
        }
        [TestMethod()]
        public void TestLotusServer()
        {
            MailSender mail = new MailSender();
            mail.SendMail(null, null);
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
            ConfigFile.ConfigFile Parameters = new ConfigFile.ConfigFile();
            OutlookAuto t = new OutlookAuto();
            t.StartMessageOit(Parameters);
        }
        [TestMethod()]
        public void FindUserLotus()
        {

          //  var Session = new NotesSession();
          ////  Session.Initialize("Password");

          //  Session.Initialize("12345");
          
          //  var registration = Session.CreateRegistration();
          //  registration.SwitchToID(@"C:\Program Files (x86)\lotus\notes\data\7751OI.id","Password");
              OutlookAuto auto = new OutlookAuto();
              var parameters = new ConfigFile.ConfigFile();
            ////  auto.StartMessageOit(parameters);
              auto.StartMessageR7751(parameters);
        }
    }
}