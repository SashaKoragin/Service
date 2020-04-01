using System.Configuration;

namespace LotusLibrary.DbConnected
{
   public class LotusConfig
    {
        public LotusConfig()
        {
            LotusServer = ConfigurationManager.AppSettings["LotusServer"];
            LotusIdFilePassword = ConfigurationManager.AppSettings["LotusIdFilePassword"];
            LotusMailSend = ConfigurationManager.AppSettings["LotusMailSend"];
            PathGenerateScheme = ConfigurationManager.AppSettings["PathGenerateScheme"];

        }

        /// <summary>
        /// Сервер Lotus
        /// </summary>
        public string LotusServer { get; set; }
        /// <summary>
        /// Пароль Lotus
        /// </summary>
        public string LotusIdFilePassword { get; set; }
        /// <summary>
        /// Почта
        /// </summary>
        public string LotusMailSend { get; set; }

        public string PathGenerateScheme { get; set; }

    }
}
