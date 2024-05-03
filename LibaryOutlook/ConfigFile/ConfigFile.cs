using System;
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
            IsSendMailOit7751 = Convert.ToBoolean(ConfigurationManager.AppSettings["IsSendMailOit7751"]);
            IsReceptionR7751 = Convert.ToBoolean(ConfigurationManager.AppSettings["IsReceptionR7751"]);
            Interval = ConfigurationManager.AppSettings["Interval"];
            //Настройки обратной связи с Консультант+
            Hours = Convert.ToInt32(ConfigurationManager.AppSettings["Hours"]);
            Minutes = Convert.ToInt32(ConfigurationManager.AppSettings["Minutes"]);
            MailReport = ConfigurationManager.AppSettings["MailReport"];
            ExtensionsFileReport = ConfigurationManager.AppSettings["ExtensionsFileReport"];
            PathConsultantPlusReceive = ConfigurationManager.AppSettings["PathConsultantPlusReceive"];
            PathConsultantPlusReceiveTemp = ConfigurationManager.AppSettings["PathConsultantPlusReceiveTemp"];
            PathConsultantPlusSts = ConfigurationManager.AppSettings["PathConsultantPlusSts"];
            IsSendReportPathConsultantPlus = Convert.ToBoolean(ConfigurationManager.AppSettings["IsSendReportPathConsultantPlus"]);
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
        /// <summary>
        /// Час запуска отправки отчета
        /// </summary>
        public int Hours { get; set; }
        /// <summary>
        /// Минута запуска отправки отчета
        /// </summary>
        public int Minutes { get; set; }
        /// <summary>
        /// Почта обратной связи с консультант+
        /// </summary>
        public string MailReport { get; set; }
        /// <summary>
        /// Файлы расширения для отчетов
        /// </summary>
        public string ExtensionsFileReport { get; set; }
        /// <summary>
        /// Путь к папке Консультант Receive
        /// </summary>
        public string PathConsultantPlusReceive { get; set; }
        /// <summary>
        /// Путь к папке Консультант Receive_Temp
        /// </summary>
        public string PathConsultantPlusReceiveTemp { get; set; }
        /// <summary>
        /// Путь к папке Консультант Sts
        /// </summary>
        public string PathConsultantPlusSts { get; set; }
        /// <summary>
        /// Отправка почты c Oit7751
        /// </summary>
        public bool IsSendMailOit7751 { get; set; }
        /// <summary>
        /// Прием почты с R7751
        /// </summary>
        public bool IsReceptionR7751 { get; set; }
        /// <summary>
        /// Отправка отчетов ConsultantPlus
        /// </summary>
        public bool IsSendReportPathConsultantPlus { get; set; }

    }
}
