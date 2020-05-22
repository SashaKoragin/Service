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
    }
}
