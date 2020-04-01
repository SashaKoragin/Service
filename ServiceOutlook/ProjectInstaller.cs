using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;


namespace ServiceOutlook
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        public ProjectInstaller()
        {
            var process = new ServiceProcessInstaller {Account = ServiceAccount.LocalSystem};
            var service = new ServiceInstaller
            {
                ServiceName = "ServiceOutlook", Description = "Служба для пересылки почты!!!"
            };
            Installers.Add(process);
            Installers.Add(service);
        }
    }
}
