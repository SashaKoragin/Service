using System;
using System.ServiceProcess;
using System.ServiceModel;
using System.Threading;
using ServiceOutlook.Service;


using LibraryOutlook.StartTaskOutlook;


namespace ServiceOutlook
{
    public partial class ServiceOutlook : ServiceBase
    {

        public StartTaskOutlook Task { get; set; }
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
            Task = new StartTaskOutlook();
            Loggers.Log4NetLogger.Info(new Exception("Запустили сервис переправки писем!"));
        }

        protected override void OnStop()
        {
            try
            {
                Task?.Dispose();
                if (Service != null)
                {
                    Service.Close();
                    Service = null;
                }
                Loggers.Log4NetLogger.Info(new Exception("Остановили сервис переправки писем!"));
            }
            catch (Exception e)
            {
                Loggers.Log4NetLogger.Info(e);
            }
        }
    }
}