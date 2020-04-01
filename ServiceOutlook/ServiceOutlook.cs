using System;
using System.ServiceProcess;
using System.ServiceModel;
using System.Threading;
using ServiceOutlook.Service;
using System.Windows.Forms;
using LibraryOutlook.Forms;



namespace ServiceOutlook
{
    public partial class ServiceOutlook : ServiceBase
    {
        public ServiceOutlook()
        {
            InitializeComponent();
        }

        public ServiceHost Service;

        protected override void OnStart(string[] args)
        {
            Service?.Close();
            Service = new ServiceHost(typeof(ServiceTest));
            Service.Open();
            new Thread(StartOutlook).Start();
        }

        void StartOutlook()
        {
            Loggers.Log4NetLogger.Info(new Exception("Запустили сервис переправки писем!"));
            Application.Run(new FormStartAutoService());
        }

        protected override void OnStop()
        {
            try
            {
                if (Service != null)
                {
                    Service.Close();
                    Service = null;
                }
                Application.Exit();
                Loggers.Log4NetLogger.Info(new Exception("Остановили сервис переправки писем!"));
            }
            catch (Exception e)
            {
                Loggers.Log4NetLogger.Info(e);
            }
        }
    }
}
