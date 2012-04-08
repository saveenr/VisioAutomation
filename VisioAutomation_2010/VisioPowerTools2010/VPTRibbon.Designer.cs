namespace VisioPowerTools2010
{
    partial class VPTRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public VPTRibbon()
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
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonImportColors = this.Factory.CreateRibbonButton();
            this.buttonCreateStencilCatalog = this.Factory.CreateRibbonButton();
            this.buttonCreateStyle = this.Factory.CreateRibbonButton();
            this.buttonHelp = this.Factory.CreateRibbonButton();
            this.groupText = this.Factory.CreateRibbonGroup();
            this.buttonToggleTextCase = this.Factory.CreateRibbonButton();
            this.buttonCopyText = this.Factory.CreateRibbonButton();
            this.buttonDeveloper = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupText.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group1);
            this.tab2.Groups.Add(this.groupText);
            this.tab2.Label = "Power Tools";
            this.tab2.Name = "tab2";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonImportColors);
            this.group1.Items.Add(this.buttonCreateStencilCatalog);
            this.group1.Items.Add(this.buttonCreateStyle);
            this.group1.Items.Add(this.buttonHelp);
            this.group1.Label = "Tools";
            this.group1.Name = "group1";
            // 
            // buttonImportColors
            // 
            this.buttonImportColors.Label = "Import colors";
            this.buttonImportColors.Name = "buttonImportColors";
            this.buttonImportColors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImportColors_Click);
            // 
            // buttonCreateStencilCatalog
            // 
            this.buttonCreateStencilCatalog.Label = "Create Stencil Catalog";
            this.buttonCreateStencilCatalog.Name = "buttonCreateStencilCatalog";
            this.buttonCreateStencilCatalog.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateStencilCatalog_Click);
            // 
            // buttonCreateStyle
            // 
            this.buttonCreateStyle.Label = "Create Style";
            this.buttonCreateStyle.Name = "buttonCreateStyle";
            this.buttonCreateStyle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateStyle_Click);
            // 
            // buttonHelp
            // 
            this.buttonHelp.Label = "Help";
            this.buttonHelp.Name = "buttonHelp";
            this.buttonHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonHelp_Click_1);
            // 
            // groupText
            // 
            this.groupText.Items.Add(this.buttonToggleTextCase);
            this.groupText.Items.Add(this.buttonCopyText);
            this.groupText.Items.Add(this.buttonDeveloper);
            this.groupText.Label = "Text";
            this.groupText.Name = "groupText";
            // 
            // buttonToggleTextCase
            // 
            this.buttonToggleTextCase.Label = "Toggle Text Case";
            this.buttonToggleTextCase.Name = "buttonToggleTextCase";
            this.buttonToggleTextCase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonToggleTextCase_Click);
            // 
            // buttonCopyText
            // 
            this.buttonCopyText.Label = "Copy text";
            this.buttonCopyText.Name = "buttonCopyText";
            this.buttonCopyText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCopyText_Click);
            // 
            // buttonDeveloper
            // 
            this.buttonDeveloper.Label = "Developer";
            this.buttonDeveloper.Name = "buttonDeveloper";
            this.buttonDeveloper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDeveloper_Click);
            // 
            // VPTRibbon
            // 
            this.Name = "VPTRibbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.VPTRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.groupText.ResumeLayout(false);
            this.groupText.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonImportColors;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateStencilCatalog;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateStyle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonToggleTextCase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCopyText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDeveloper;
    }

    partial class ThisRibbonCollection
    {
        internal VPTRibbon VPTRibbon
        {
            get { return this.GetRibbon<VPTRibbon>(); }
        }
    }
}
