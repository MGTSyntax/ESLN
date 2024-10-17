namespace ESLNotification
{
    partial class eslnProjectInstaller
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
            this.eslnProcessInstaller = new System.ServiceProcess.ServiceProcessInstaller();
            this.eslnInstaller = new System.ServiceProcess.ServiceInstaller();
            // 
            // eslnProcessInstaller
            // 
            this.eslnProcessInstaller.Account = System.ServiceProcess.ServiceAccount.LocalSystem;
            this.eslnProcessInstaller.Password = null;
            this.eslnProcessInstaller.Username = null;
            // 
            // eslnInstaller
            // 
            this.eslnInstaller.Description = "Notification for Expiry Security License";
            this.eslnInstaller.DisplayName = "ESLN";
            this.eslnInstaller.ServiceName = "ESLN";
            this.eslnInstaller.StartType = System.ServiceProcess.ServiceStartMode.Automatic;
            // 
            // eslnProjectInstaller
            // 
            this.Installers.AddRange(new System.Configuration.Install.Installer[] {
            this.eslnProcessInstaller,
            this.eslnInstaller});

        }

        #endregion

        private System.ServiceProcess.ServiceProcessInstaller eslnProcessInstaller;
        private System.ServiceProcess.ServiceInstaller eslnInstaller;
    }
}