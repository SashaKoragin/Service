using System;
using System.Runtime.InteropServices;
using Domino;

namespace LotusLibrary.DbConnected
{
   public class LotusConnectedDataBase:IDisposable
    {
        public NotesSession Session { get; set; }
        public NotesDatabase Db { get; set; }
        /// <summary>
        /// Document
        /// </summary>
        public NotesDocument Document { get; set; }
        /// <summary>
        /// View
        /// </summary>
        public NotesView NotesView { get; set; }
        /// <summary>
        /// Наименование БД для log
        /// </summary>
        public string NameDateBase { get; set; }
        /// <summary>
        /// Инициализация Пароль
        /// </summary>
        /// <param name="password">Пароль</param>
        public LotusConnectedDataBase(string password)
        {
            Dispose();
            Session = new NotesSession();
            Session.Initialize(password);
        }
        /// <summary>
        /// Переключение на Id Файл
        /// </summary>
        /// <param name="idPath">Путь к Id Файлу</param>
        /// <param name="password">Пароль</param>
        public void SwitchIdUser(string idPath, string password)
        {
            var registration = Session.CreateRegistration();
            registration.SwitchToID(idPath, password);
        }

        /// <summary>
        /// Подключение к БД
        /// </summary>
        /// <param name="server">Сервер</param>
        /// <param name="database">Наименование БД</param>
        /// <param name="isCreateDb">Создаем БД или нет по умолчанию нет</param>
        public NotesDatabase LotusConnectedDataBaseServer(string server, string database, bool isCreateDb = false)
        {
            try
            {
                Loggers.Log4NetLogger.Info(new Exception($"Пользователь {Session.UserName}"));
                Db = Session.GetDatabase(server, database, isCreateDb);
                if (Db == null)
                    throw new InvalidOperationException("Фатальная ошибка нет соединения с сервером!");
                if (Db.IsOpen)
                {
                    NameDateBase = Db.FileName;
                    Loggers.Log4NetLogger.Info(new Exception($"База данных {NameDateBase} открыта!"));
                    return Db;
                };
                Loggers.Log4NetLogger.Error(new Exception($"База данных {Db.FileName} не открыта!"));
                return null;
            }
            catch(Exception ex)
            {
                Db = null;
                Loggers.Log4NetLogger.Error(new Exception($"Соединение с базой данных {database} не возможно!"));
                Loggers.Log4NetLogger.Error(ex);
            }
            return null;
        }
        /// <summary>
        /// Удаление всей почты если переполнение документами и сжатие БД
        /// Нет доступа на Compact() БД
        /// </summary>
        public void DeleteDataBaseAllMailSizeWarning()
        {
            var sizeWarning = Db.SizeWarning * 0.001; //Мегабайт
            var size = Db.Size/1024/1024;  //Байты в Мегабайты
            if (size > sizeWarning)
            {
                try
                {
                    Db.AllDocuments.RemoveAll(true);
                    NotesView = GetViewLotus("$SoftDeletions");
                    Document = NotesView.GetFirstDocument();
                    while (Document != null)
                    {
                       var universalId = Document.UniversalID;
                       var doc = Document;
                       Document = NotesView.GetNextDocument(Document);
                       doc.RemovePermanently(true);
                       if (doc.IsDeleted)
                       {
                          Loggers.Log4NetLogger.Info(new Exception($"Документ под ID: {universalId} удален!")); 
                       }
                    }
                    NotesView.Refresh();
                    Loggers.Log4NetLogger.Error(new Exception($"Срочно требуется сжатие БД {Db.FileName}!!!"));
                }
                catch (Exception ex)
                {
                    Loggers.Log4NetLogger.Error(ex);
                }
                finally
                {
                    if (Document != null)
                        Marshal.ReleaseComObject(Document);
                    Document = null;
                    if (NotesView != null)
                        Marshal.ReleaseComObject(NotesView);
                    NotesView = null;
                }
            }
        }



        /// <summary>
        /// Получение представления по имени
        /// </summary>
        /// <param name="nameView">Наименование представления</param>
        /// <returns></returns>
        public NotesView GetViewLotus(string nameView)
        {
            return Db.GetView(nameView);
        }

        /// <summary>
        /// Стиль текста
        /// </summary>
        public NotesRichTextStyle StyleStandartText()
        {
           var styleText = Session.CreateRichTextStyle();
           styleText.FontSize = 16;
           styleText.NotesColor = COLORS.COLOR_BLACK;
           return styleText;
        }


        /// <summary>
        /// Индексирование БД если индекса нет то делаем
        /// </summary>
        /// <param name="dbDatabase"></param>
        private NotesDatabase IndexDataBaseUseCreate(NotesDatabase dbDatabase)
        {
            if (dbDatabase.IsFTIndexed)
            {
                return dbDatabase;
            }
            dbDatabase.CreateFTIndex(FTINDEX_OPTIONS.FTINDEX_ALL_BREAKS & FTINDEX_OPTIONS.FTINDEX_CASE_SENSITIVE, false);
            return dbDatabase;
        }

        private void ReleaseUnmanagedResources()
        {
            if (Db != null)
                Marshal.ReleaseComObject(Db);
            Db = null;
            if (Session != null)
                Marshal.ReleaseComObject(Session);
            Session = null;
            if(NameDateBase!=null)
                Loggers.Log4NetLogger.Info(new Exception($"База данных {NameDateBase} сброшена из ресурса процессов!"));
        }

        protected virtual void Dispose(bool disposing)
        {
            ReleaseUnmanagedResources();
            if (disposing)
            {
            }
        }

        public void Dispose()
        {
            Dispose(true);
           // GC.SuppressFinalize(this);
        }

        ~LotusConnectedDataBase()
        {
            Dispose(false);
        }
    }
}
