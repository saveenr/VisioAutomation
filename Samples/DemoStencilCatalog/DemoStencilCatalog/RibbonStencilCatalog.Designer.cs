namespace DemoStencilCatalog
{
    partial class RibbonStencilCatalog : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonStencilCatalog()
            : base(Globals.Factory.GetRibbonFactory())
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupStencilCatalog = this.Factory.CreateRibbonGroup();
            this.buttonRun = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupStencilCatalog.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupStencilCatalog);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupStencilCatalog
            // 
            this.groupStencilCatalog.Items.Add(this.buttonRun);
            this.groupStencilCatalog.Label = "StencilCatalog";
            this.groupStencilCatalog.Name = "groupStencilCatalog";
            // 
            // buttonRun
            // 
            this.buttonRun.Label = "Create Catalog";
            this.buttonRun.Name = "buttonRun";
            this.buttonRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRun_Click);
            // 
            // RibbonStencilCatalog
            // 
            this.Name = "RibbonStencilCatalog";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonStencilCatalog_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupStencilCatalog.ResumeLayout(false);
            this.groupStencilCatalog.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupStencilCatalog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRun;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonStencilCatalog RibbonStencilCatalog
        {
            get { return this.GetRibbon<RibbonStencilCatalog>(); }
        }
    }
}
