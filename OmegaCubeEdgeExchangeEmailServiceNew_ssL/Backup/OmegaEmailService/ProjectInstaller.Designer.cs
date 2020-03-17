namespace OmegaEmailService
{
    partial class ProjectInstaller
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.OmegaEmailserviceProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.OmegaEmailserviceInstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // OmegaEmailserviceProcessInstaller
            // 
            this.OmegaEmailserviceProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.OmegaEmailserviceProcessInstaller.Password = null;
            this.OmegaEmailserviceProcessInstaller.Username = null;
            this.OmegaEmailserviceProcessInstaller.AfterInstall += new System.Configuration.Install.InstallEventHandler(this.serviceProcessInstaller1_AfterInstall);
            // 
            // OmegaEmailserviceInstaller
            // 
            this.OmegaEmailserviceInstaller.ServiceName = "OmegacubeEmailService";
            this.OmegaEmailserviceInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            this.OmegaEmailserviceInstaller.BeforeUninstall += new System.Configuration.Install.InstallEventHandler(this.OmegaEmailserviceInstaller_BeforeUninstall);
            this.OmegaEmailserviceInstaller.AfterUninstall += new System.Configuration.Install.InstallEventHandler(this.OmegaEmailserviceInstaller_AfterUninstall);
            // 
            // ProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.OmegaEmailserviceProcessInstaller,
            this.OmegaEmailserviceInstaller});
            this.AfterUninstall += new System.Configuration.Install.InstallEventHandler(this.ProjectInstaller_AfterUninstall);

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller OmegaEmailserviceProcessInstaller;
        private System.ServiceProcess.ServiceInstaller OmegaEmailserviceInstaller;
    }
}