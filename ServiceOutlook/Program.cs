using System.ServiceProcess;


namespace ServiceOutlook
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        static void Main()
        {
            var serviceRun = new ServiceBase[]
            {
                new ServiceOutlook() 
            };
            ServiceBase.Run(serviceRun);
        }
    }
}
