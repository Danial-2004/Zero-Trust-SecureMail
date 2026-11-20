using SecureMailAddin;
using System;
using System.Numerics;

namespace SecureOutlookAddin
{
    partial class SecureMailRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public SecureMailRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        private void InitializeComponent()
        {
            this.SecureMailTab = this.Factory.CreateRibbonTab();
            this.SecurityGroup = this.Factory.CreateRibbonGroup();
            this.EncryptButton = this.Factory.CreateRibbonButton();
            this.DecryptButton = this.Factory.CreateRibbonButton();
            this.ExportKeyButton = this.Factory.CreateRibbonButton();
            this.ImportKeyButton = this.Factory.CreateRibbonButton();
            this.SecureMailTab.SuspendLayout();
            this.SecurityGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // SecureMailTab
            // 
            this.SecureMailTab.Groups.Add(this.SecurityGroup);
            this.SecureMailTab.Label = "Secure Mail";
            this.SecureMailTab.Name = "SecureMailTab";
            // 
            // SecurityGroup
            // 
            this.SecurityGroup.Items.Add(this.EncryptButton);
            this.SecurityGroup.Items.Add(this.DecryptButton);
            this.SecurityGroup.Items.Add(this.ExportKeyButton);
            this.SecurityGroup.Items.Add(this.ImportKeyButton);
            this.SecurityGroup.Label = "Security";
            this.SecurityGroup.Name = "SecurityGroup";
            // 
            // EncryptButton
            // 
            this.EncryptButton.Label = "Encrypt";
            this.EncryptButton.Name = "EncryptButton";
            this.EncryptButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EncryptButton_Click);
            // 
            // DecryptButton
            // 
            this.DecryptButton.Label = "Decrypt";
            this.DecryptButton.Name = "DecryptButton";
            this.DecryptButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DecryptButton_Click);
            // 
            // ExportKeyButton
            // 
            this.ExportKeyButton.Label = "Export Key";
            this.ExportKeyButton.Name = "ExportKeyButton";
            this.ExportKeyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportKeyButton_Click);
            // 
            // ImportKeyButton
            // 
            this.ImportKeyButton.Label = "Import Key";
            this.ImportKeyButton.Name = "ImportKeyButton";
            this.ImportKeyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ImportKeyButton_Click);
            // 
            // SecureMailRibbon
            // 
            this.Name = "SecureMailRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer,Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.SecureMailTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SecureMailRibbon_Load);
            this.SecureMailTab.ResumeLayout(false);
            this.SecureMailTab.PerformLayout();
            this.SecurityGroup.ResumeLayout(false);
            this.SecurityGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab SecureMailTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup SecurityGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EncryptButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DecryptButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportKeyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ImportKeyButton;

    }
}
