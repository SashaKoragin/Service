using System;
using System.Timers;
using LibraryOutlook.SubscribeOutlook;

namespace LibraryOutlook.StartTaskOutlook
{
   public class StartTaskOutlook : IDisposable
    {
        private OutlookAutoPop3 OutlookAutoPop3 { get; set; }

        private OutlookAutoSmtp OutlookAutoSmtp { get; set; }

        private readonly ConfigFile.ConfigFile parameters = new ConfigFile.ConfigFile();

        private readonly Timer timerMessage = new Timer();

        /// <summary>
        /// Класс запуска фоновой задачи
        /// </summary>
        public StartTaskOutlook()
        {
            OutlookAutoPop3 = new OutlookAutoPop3();
            OutlookAutoSmtp = new OutlookAutoSmtp();
            timerMessage.Interval = Convert.ToInt32(parameters.Interval);
            timerMessage.Enabled = true;
            timerMessage.AutoReset = true;
            timerMessage.Elapsed += TimerMessage_Tick;
            timerMessage.Start();
        }

        /// <summary>
        /// Мониторинг на 2 почты
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerMessage_Tick(object sender, EventArgs e)
        {
            OutlookAutoPop3.StartMessageOit(parameters);
            OutlookAutoPop3.StartMessageR7751(parameters);
            OutlookAutoSmtp.SendSmtpMessage(parameters);
        }
        /// <summary>
        /// Освобождение памяти
        /// </summary>
        public void Dispose()
        {
            OutlookAutoPop3 = null;
            OutlookAutoSmtp = null;
            timerMessage.Dispose();
        }
    }
}
