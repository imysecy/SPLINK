namespace SharePoint_Link
{
    partial class SharePointRibbon : Microsoft.Office.Tools.Ribbon.OfficeRibbon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SharePointRibbon()
        {
            InitializeComponent();
        }

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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncher1 = new Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SharePointRibbon));
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.group1 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnConnectionProperties = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.btnNewConnection = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.btnOptions = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.button1 = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "ITOPIA  ";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncher1;
            this.group1.Items.Add(this.btnConnectionProperties);
            this.group1.Items.Add(this.btnNewConnection);
            this.group1.Items.Add(this.btnOptions);
            this.group1.Label = "ITOPIA SharePoint Configuration";
            this.group1.Name = "group1";
            // 
            // btnConnectionProperties
            // 
            this.btnConnectionProperties.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConnectionProperties.Description = "Connect To SharePoint";
            this.btnConnectionProperties.Image = global::SharePoint_Link.Properties.Resources.conn_new;
            this.btnConnectionProperties.ImageName = "Connection Properties";
            this.btnConnectionProperties.Label = "Connection Properties";
            this.btnConnectionProperties.Name = "btnConnectionProperties";
            this.btnConnectionProperties.ScreenTip = "Connection Properties";
            this.btnConnectionProperties.ShowImage = true;
            this.btnConnectionProperties.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnConnectionProperties_Click);
            // 
            // btnNewConnection
            // 
            this.btnNewConnection.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewConnection.Description = "Connect To SharePoint";
            this.btnNewConnection.Image = global::SharePoint_Link.Properties.Resources.new_conn;
            this.btnNewConnection.ImageName = "New Connection ";
            this.btnNewConnection.Label = "New Connection";
            this.btnNewConnection.Name = "btnNewConnection";
            this.btnNewConnection.ScreenTip = "New Connection ";
            this.btnNewConnection.ShowImage = true;
            this.btnNewConnection.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnNewConnection_Click);
            // 
            // btnOptions
            // 
            this.btnOptions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOptions.Description = "Connect To SharePoint";
            this.btnOptions.Image = global::SharePoint_Link.Properties.Resources.option;
            this.btnOptions.ImageName = "Options";
            this.btnOptions.Label = "Options";
            this.btnOptions.Name = "btnOptions";
            this.btnOptions.ScreenTip = "Options";
            this.btnOptions.ShowImage = true;
            this.btnOptions.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnOptions_Click);
            // 
            // button1
            // 
            this.button1.Label = "SSP 2010";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            // 
            // SharePointRibbon
            // 
            this.Name = "SharePointRibbon";
            // 
            // SharePointRibbon.OfficeMenu
            // 
            this.OfficeMenu.Items.Add(this.button1);
            this.RibbonType = resources.GetString("$this.RibbonType");
            this.Tabs.Add(this.tab1);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.SharePointRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnectionProperties;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewConnection;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOptions;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal SharePointRibbon SharePointRibbon
        {
            get { return this.GetRibbon<SharePointRibbon>(); }
        }
    }
}
