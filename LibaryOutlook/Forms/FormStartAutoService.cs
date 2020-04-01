using System;
using System.Windows.Forms;
using LibraryOutlook.SubscribeOutlook;
using LotusLibrary.DbConnected;
using LotusLibrary.MailSender;

namespace LibraryOutlook.Forms
{
    public class FormStartAutoService : Form
    {
        public ConfigFile.ConfigFile Parameters { get; set; }
        public OutlookAuto OutlookAuto {get; set; }

        public Timer TimerMessage = new Timer(); 

        public FormStartAutoService()
        {
            SuspendLayout();
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(0, 0);
            FormBorderStyle = FormBorderStyle.None;
            Name = $"Скрытая форма службы Outlook!";
            Text = "Скрытая форма службы Outlook!";
            WindowState = FormWindowState.Minimized;
            Load += HiddenForm_Load;
            FormClosing += HiddenForm_FormClosing;
            Parameters = new ConfigFile.ConfigFile();
            TimerMessage.Interval = Convert.ToInt32(Parameters.Interval);
            ResumeLayout(false);
        }

        public sealed override string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        private void HiddenForm_Load(object sender, EventArgs e)
        {
            OutlookAuto = new OutlookAuto();
            TimerMessage.Tick += TimerMessage_Tick;
            TimerMessage.Start();
        }
        /// <summary>
        /// Мониторинг на 2 почты
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerMessage_Tick(object sender, EventArgs e)
        {
            OutlookAuto.StartMessageOit(Parameters);
            OutlookAuto.StartMessageR7751(Parameters);
        }

        private void HiddenForm_FormClosing(object sender, FormClosingEventArgs e)
        {
        }
    }
}
