using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Domino;

namespace LotusLibrary.DxlLotus.DocumentGeneration
{
   public class DonloadOnCreateDxlFile
    {
        /// <summary>
        /// Загрузка файла
        /// </summary>
        /// <param name="fileName">Путь к файлу</param>
        /// <param name="db">БД</param>
        public string ImportDxlFile(string fileName, NotesDatabase db)
        {
            NotesStream notesStream = null;
            NotesDXLImporter notesDXLImporter = null;
            notesStream = db.Parent.CreateStream();
            notesDXLImporter = db.Parent.CreateDXLImporter();
            if (!notesStream.Open(fileName))
            {
                Loggers.Log4NetLogger.Error(new Exception("Невозможно открыть файл " + fileName));
                return null;
            }
            notesDXLImporter.ACLImportOption = DXLIMPORTOPTION.DXLIMPORTOPTION_UPDATE_ELSE_IGNORE;
            notesDXLImporter.DesignImportOption = DXLIMPORTOPTION.DXLIMPORTOPTION_REPLACE_ELSE_CREATE;
            notesDXLImporter.ReplicaRequiredForReplaceOrUpdate = false;
            notesDXLImporter.DocumentImportOption = DXLIMPORTOPTION.DXLIMPORTOPTION_UPDATE_ELSE_CREATE;
            notesDXLImporter.ExitOnFirstFatalError = true;
            try
            {
                notesDXLImporter.Import(notesStream, db);
                string text = notesDXLImporter.GetFirstImportedNoteId();
                notesStream.Truncate();
                notesStream.Close();
                return text;
            }
            catch (Exception e)
            {
                Loggers.Log4NetLogger.Error(new Exception(notesDXLImporter.Log));
                Loggers.Log4NetLogger.Error(new Exception(notesDXLImporter.LogComment));
                Loggers.Log4NetLogger.Error(e);
            }
            finally
            {
                if (notesDXLImporter != null)
                    Marshal.ReleaseComObject(notesDXLImporter);
                if (notesStream != null)
                    Marshal.ReleaseComObject(notesStream);
            }
            return null;
        }


        /// <summary>
        /// Создание DXL файла для NS:http://www.lotus.com/dxl
        /// </summary>
        /// <param name="pathName">Путь к файлу file.dxl</param>
        /// <param name="classXml">Класс схемы</param>
        /// <param name="typeXml">Тип Xml</param>
        public void DxlFileSave(string pathName,object classXml, Type typeXml )
        {
            XmlSerializerNamespaces ns = new XmlSerializerNamespaces();
            ns.Add(string.Empty, "http://www.lotus.com/dxl");
            string str;
            using (StringWriter textWriter = new StringWriter())
            {
                XmlSerializer formatter = new XmlSerializer(typeXml);
                formatter.Serialize(textWriter, classXml, ns);
                str = textWriter.ToString();
            }
            using (FileStream output = new FileStream(pathName, FileMode.Create, FileAccess.Write, FileShare.Delete))
            {
                XmlDocument xmlDocument = new XmlDocument();
                XmlWriterSettings settings = new XmlWriterSettings
                {
                    Indent = true,
                    IndentChars = "  ",
                    Encoding = Encoding.UTF8,
                    NewLineChars = "\r\n",
                    NewLineHandling = NewLineHandling.Replace
                };
                using (XmlWriter w = XmlWriter.Create(output, settings))
                {
                    xmlDocument.LoadXml(str);
                    xmlDocument.Save(w);
                }
            }
        }

    }
}
