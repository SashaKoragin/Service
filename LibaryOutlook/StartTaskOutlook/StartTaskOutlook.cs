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
        /// Временная задача для отправки отчета по Консультант +
        /// </summary>
        private readonly Timer timerConsultantPlusSendReport = new Timer();
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
            timerConsultantPlusSendReport.Interval = 60000;
            timerConsultantPlusSendReport.Enabled = true;
            timerConsultantPlusSendReport.AutoReset = true;
            timerConsultantPlusSendReport.Elapsed += TimerSendReportConsultantPlus;
            timerConsultantPlusSendReport.Start();

        }

        /// <summary>
        /// Мониторинг на 2 почты
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerMessage_Tick(object sender, EventArgs e)
        {
            OutlookAutoPop3.StartMessageOit(parameters);
            if (parameters.IsSendMailMy)
            {
                OutlookAutoSmtp.SendSmtpMessageOit(parameters);
            }
            if (!parameters.IsReceptionR7751) return;
            OutlookAutoPop3.StartMessageR7751(parameters);
        }
        /// <summary>
        /// Отправка отчетов в офис консультанта плюс
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerSendReportConsultantPlus(object sender, EventArgs e)
        {
            DateTime date = DateTime.Now;
            if (date.Hour == parameters.Hours && date.Minute == parameters.Minutes && parameters.IsSendReportPathConsultantPlus)
            {
                OutlookAutoSmtp.SendSmtpConsultantPlusReport(parameters);
            }
        }

        /// <summary>
        /// Освобождение памяти
        /// </summary>
        public void Dispose()
        {
            OutlookAutoPop3 = null;
            OutlookAutoSmtp = null;
            timerMessage.Dispose();
            timerConsultantPlusSendReport.Dispose();
        }
    }
}
