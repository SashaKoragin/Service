using System;
using System.Timers;
using LibraryOutlook.SubscribeOutlook;

namespace LibraryOutlook.StartTaskOutlook
{
   public class StartTaskOutlook : IDisposable
    {
        private OutlookAuto OutlookAuto { get; set; }

        private readonly ConfigFile.ConfigFile parameters = new ConfigFile.ConfigFile();

        private readonly Timer timerMessage = new Timer();

        /// <summary>
        /// Класс запуска фоновой задачи
        /// </summary>
        public StartTaskOutlook()
        {
            OutlookAuto = new OutlookAuto();
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
            OutlookAuto.StartMessageOit(parameters);
            OutlookAuto.StartMessageR7751(parameters);
        }
        /// <summary>
        /// Освобождение памяти
        /// </summary>
        public void Dispose()
        {
            OutlookAuto = null;
            timerMessage.Dispose();
        }
    }
}
