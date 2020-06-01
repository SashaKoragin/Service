using System.Collections.Generic;
using System.IO;
using Domino;

namespace LotusLibrary.PublicFunctionInfo
{
   public class PublicFunctionInfo
    {
        /// <summary>
        /// Запись всех представлений в лог
        /// </summary>
        public void AllViewToLog(NotesDatabase db)
        {
            using (StreamWriter outputFile = new StreamWriter(Path.Combine(@"C:\Users\7751_svc_admin\Desktop\", "WriteLines.txt")))
            {
                foreach (NotesView dbDatabaseView in db.Views)
                {
                    outputFile.WriteLine($"Поля и текст {dbDatabaseView.Name}");
                }
            }
        }

        /// <summary>
        /// Извлечение файла в указанный путь
        /// </summary>
        /// <param name="notesRich">Элемент notesRich</param>
        /// <param name="path">Путь к файлу</param>
        /// <returns>Возврат всех наименований файлов вложений</returns>
        public List<string> ExtractFile(NotesRichTextItem notesRich, string path)
        {
            if (notesRich != null)
            {
                if (!string.IsNullOrWhiteSpace(notesRich.Values))
                {
                    if (notesRich.type == IT_TYPE.RICHTEXT)
                    {
                        List<string> listFullPath = new List<string>();
                        if (notesRich.EmbeddedObjects != null)
                        {
                            foreach (var embedded in notesRich.EmbeddedObjects)
                            {
                                if (embedded.Type == 1454)
                                {
                                    var fileName = path + embedded.Name;
                                    embedded.ExtractFile(fileName);
                                    listFullPath.Add(fileName);
                                }
                            }
                            return listFullPath;
                        }
                    }
                }
            }
            return null;
        }
    }
}
