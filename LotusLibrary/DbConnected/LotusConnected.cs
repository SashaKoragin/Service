using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Domino;

namespace LotusLibrary.DbConnected
{
   public class LotusConnectedDataBase:IDisposable
    {
        public NotesSession Session { get; set; }
        public NotesDatabase Db { get; set; }
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
        public NotesDatabase LotusConnectedDataBaseServer( string server, string database, bool isCreateDb = false)
        {
            try
            {
                Loggers.Log4NetLogger.Info(new Exception($"Пользователь {Session.UserName}"));
                Db = Session.GetDatabase(server, database, isCreateDb);
                if (Db.IsOpen)
                {
                    Loggers.Log4NetLogger.Info(new Exception($"База данных {Db.FileName} открыта!"));
                    return Db;
                };
                Loggers.Log4NetLogger.Error(new Exception($"База данных {Db.FileName} не открыта!"));
                return null;
            }
            catch(Exception ex)
            {
                Loggers.Log4NetLogger.Error(ex);
            }
            return null;
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

        public void Dispose()
        {
            if (Db != null)
                Marshal.ReleaseComObject(Db);
            Db = null;
            if (Session!= null)
                Marshal.ReleaseComObject(Session);
            Session = null;
        }
    }
}
