using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Domino;

namespace LotusLibrary.DxlLotus.DocumentGeneration
{
   public class DonloadOnCreateDxlFile:IDisposable
    {
        public NotesStream NotesStream { get; set; }

        public NotesDXLImporter NotesDxlImporter { get; set; }

        public DonloadOnCreateDxlFile()
        {
            Dispose();
        }

        /// <summary>
        /// Загрузка файла
        /// </summary>
        /// <param name="fileName">Путь к файлу</param>
        /// <param name="db">БД</param>
        public string ImportDxlFile(string fileName, NotesDatabase db)
        {
            NotesStream = db.Parent.CreateStream();
            NotesDxlImporter = db.Parent.CreateDXLImporter();
            if (!NotesStream.Open(fileName))
            {
                Loggers.Log4NetLogger.Error(new Exception("Невозможно открыть файл " + fileName));
                return null;
            }
            //notesDXLImporter.InputValidationOption = VALIDATIONOPTION.VALIDATE_NEVER;
            NotesDxlImporter.ACLImportOption = DXLIMPORTOPTION.DXLIMPORTOPTION_UPDATE_ELSE_IGNORE;
            NotesDxlImporter.DesignImportOption = DXLIMPORTOPTION.DXLIMPORTOPTION_REPLACE_ELSE_CREATE;
            NotesDxlImporter.ReplicaRequiredForReplaceOrUpdate = false;
            NotesDxlImporter.DocumentImportOption = DXLIMPORTOPTION.DXLIMPORTOPTION_UPDATE_ELSE_CREATE;
            NotesDxlImporter.ExitOnFirstFatalError = true;
            try
            {
                NotesDxlImporter.Import(NotesStream, db);
                string text = NotesDxlImporter.GetFirstImportedNoteId();
                NotesStream.Truncate();
                NotesStream.Close();
                return text;
            }
            catch (Exception e)
            {
                Loggers.Log4NetLogger.Error(new Exception(NotesDxlImporter.Log));
                Loggers.Log4NetLogger.Error(new Exception(NotesDxlImporter.LogComment));
                Loggers.Log4NetLogger.Error(e);
            }
            finally
            {
                Dispose();
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
        /// <summary>
        /// Освобождение ресурсов памяти
        /// </summary>
        public void Dispose()
        {
            if (NotesStream != null)
                Marshal.ReleaseComObject(NotesStream);
            NotesStream = null;
            if (NotesDxlImporter != null)
                Marshal.ReleaseComObject(NotesDxlImporter);
            NotesDxlImporter = null;
        }
    }
}
