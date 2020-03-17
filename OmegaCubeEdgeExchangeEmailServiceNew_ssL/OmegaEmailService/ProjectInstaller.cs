using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.ServiceProcess;
using System.IO;

namespace OmegaEmailService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        protected static string AppPath = System.AppDomain.CurrentDomain.BaseDirectory;
        private System.Diagnostics.EventLog eventLogFile = null;

        //public static string source = "OmegacubeEmailSource";
        //public static string log = "OmegacubeEmailLog";

        public ProjectInstaller()
        {
            InitializeComponent();
        }

        private void serviceProcessInstaller1_AfterInstall(object sender, InstallEventArgs e)
        {

        }

        private void ProjectInstaller_AfterUninstall(object sender, InstallEventArgs e)
        {
           
        }

        private void OmegaEmailserviceInstaller_AfterUninstall(object sender, InstallEventArgs e)
        {
            // Remove Event Source if already there    
            if (EventLog.SourceExists("OmegacubeEmailSource"))
            {
                EventLog.DeleteEventSource("OmegacubeEmailSource");
            }

            if (File.Exists(AppPath + "\\" + "OmegacubeEmailLog.txt"))
            {
                File.Delete(AppPath + "\\" + "OmegacubeEmailLog.txt");
            }
        }

        private void OmegaEmailserviceInstaller_BeforeUninstall(object sender, InstallEventArgs e)
        {
            try
            {
                ServiceController controller = new ServiceController(OmegaEmailserviceInstaller.ServiceName);
                if (controller.Status == ServiceControllerStatus.Running | controller.Status == ServiceControllerStatus.Paused)
                {
                    controller.Stop();
                    controller.WaitForStatus(ServiceControllerStatus.Stopped, new TimeSpan(0, 0, 0, 15));
                    controller.Close();
                }
            }
            catch (Exception ex)
            {
                eventLogFile.WriteEntry(ex.Message.ToString());
            }
        }
    }
}