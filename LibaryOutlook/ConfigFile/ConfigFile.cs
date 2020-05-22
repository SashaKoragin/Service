using System.Configuration;

namespace LibraryOutlook.ConfigFile
{
   public class ConfigFile
    {

        public ConfigFile()
        {
            Pop3Address = ConfigurationManager.AppSettings["Pop3Address"];
            LoginOit = ConfigurationManager.AppSettings["LoginOit"];
            PasswordOit = ConfigurationManager.AppSettings["PasswordOit"];
            PathSaveArchive = ConfigurationManager.AppSettings["PathSaveArchive"];
            LoginR7751 = ConfigurationManager.AppSettings["LoginR7751"];
            PasswordR7751 = ConfigurationManager.AppSettings["PasswordR7751"];
            Interval = ConfigurationManager.AppSettings["Interval"];

        }
        /// <summary>
        /// Интервал 10 минут
        /// </summary>
        public string Interval { get; set; }

        /// <summary>
        /// Адрес сервера POP3
        /// </summary>
        public string Pop3Address { get; set; }
        /// <summary>
        /// Логин
        /// </summary>
        public string LoginOit { get; set; }
        /// <summary>
        /// Пароль
        /// </summary>
        public string PasswordOit { get; set; }
        /// <summary>
        /// Почта общая Login
        /// </summary>
        public string LoginR7751 { get; set; }
        /// <summary>
        /// Почта общая Пароль
        /// </summary>
        public string PasswordR7751 { get; set; }
        /// <summary>
        /// Путь к сохранению архива
        /// </summary>
        public string PathSaveArchive { get; set; }
    }
}
