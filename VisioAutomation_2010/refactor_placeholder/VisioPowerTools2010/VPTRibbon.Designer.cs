using System.ComponentModel;
using Microsoft.Office.Tools.Ribbon;

namespace VisioPowerTools2010
{
    partial class VPTRibbon : RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

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
            this.groupDraw = this.Factory.CreateRibbonGroup();
            this.buttonGraph = this.Factory.CreateRibbonButton();
            this.groupExport = this.Factory.CreateRibbonGroup();
            this.buttonExportSelection = this.Factory.CreateRibbonButton();
            this.buttonSelectionXHTML = this.Factory.CreateRibbonButton();
            this.buttomResetPageOrigin = this.Factory.CreateRibbonButton();
            this.groupDev = this.Factory.CreateRibbonGroup();
            this.buttonDeveloper = this.Factory.CreateRibbonButton();
            this.buttonScrambleText = this.Factory.CreateRibbonButton();
            this.groupPage = this.Factory.CreateRibbonGroup();
            this.buttonResizePageToFit = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            this.groupText.SuspendLayout();
            this.groupDraw.SuspendLayout();
            this.groupExport.SuspendLayout();
            this.groupDev.SuspendLayout();
            this.groupPage.SuspendLayout();
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
            this.tab2.Groups.Add(this.groupPage);
            this.tab2.Groups.Add(this.groupDraw);
            this.tab2.Groups.Add(this.groupExport);
            this.tab2.Groups.Add(this.groupDev);
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
            // groupDraw
            // 
            this.groupDraw.Items.Add(this.buttonGraph);
            this.groupDraw.Label = "Draw";
            this.groupDraw.Name = "groupDraw";
            // 
            // buttonGraph
            // 
            this.buttonGraph.Label = "Directed Graph";
            this.buttonGraph.Name = "buttonGraph";
            this.buttonGraph.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGraph_Click);
            // 
            // groupExport
            // 
            this.groupExport.Items.Add(this.buttonExportSelection);
            this.groupExport.Items.Add(this.buttonSelectionXHTML);
            this.groupExport.Label = "Export";
            this.groupExport.Name = "groupExport";
            // 
            // buttonExportSelection
            // 
            this.buttonExportSelection.Label = "Selection > XAML";
            this.buttonExportSelection.Name = "buttonExportSelection";
            this.buttonExportSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonExportSelection_Click);
            // 
            // buttonSelectionXHTML
            // 
            this.buttonSelectionXHTML.Label = "Selection > XHTML";
            this.buttonSelectionXHTML.Name = "buttonSelectionXHTML";
            this.buttonSelectionXHTML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSelectionXHTML_Click);
            // 
            // buttomResetPageOrigin
            // 
            this.buttomResetPageOrigin.Label = "Reset Page Origin";
            this.buttomResetPageOrigin.Name = "buttomResetPageOrigin";
            this.buttomResetPageOrigin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttomResetPageOrigin_Click);
            // 
            // groupDev
            // 
            this.groupDev.Items.Add(this.buttonDeveloper);
            this.groupDev.Items.Add(this.buttonScrambleText);
            this.groupDev.Label = "Developer";
            this.groupDev.Name = "groupDev";
            // 
            // buttonDeveloper
            // 
            this.buttonDeveloper.Label = "Developer";
            this.buttonDeveloper.Name = "buttonDeveloper";
            this.buttonDeveloper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDeveloper_Click);
            // 
            // buttonScrambleText
            // 
            this.buttonScrambleText.Label = "ScrambleText";
            this.buttonScrambleText.Name = "buttonScrambleText";
            this.buttonScrambleText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonScrambleText_Click);
            // 
            // groupPage
            // 
            this.groupPage.Items.Add(this.buttomResetPageOrigin);
            this.groupPage.Items.Add(this.buttonResizePageToFit);
            this.groupPage.Label = "Page";
            this.groupPage.Name = "groupPage";
            // 
            // buttonResizePageToFit
            // 
            this.buttonResizePageToFit.Label = "Resize to Fit";
            this.buttonResizePageToFit.Name = "buttonResizePageToFit";
            this.buttonResizePageToFit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonResizePageToFit_Click);
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
            this.groupDraw.ResumeLayout(false);
            this.groupDraw.PerformLayout();
            this.groupExport.ResumeLayout(false);
            this.groupExport.PerformLayout();
            this.groupDev.ResumeLayout(false);
            this.groupDev.PerformLayout();
            this.groupPage.ResumeLayout(false);
            this.groupPage.PerformLayout();

        }

        #endregion

        internal RibbonTab tab1;
        private RibbonTab tab2;
        internal RibbonGroup group1;
        internal RibbonButton buttonHelp;
        internal RibbonButton buttonImportColors;
        internal RibbonButton buttonCreateStencilCatalog;
        internal RibbonButton buttonCreateStyle;
        internal RibbonGroup groupText;
        internal RibbonButton buttonToggleTextCase;
        internal RibbonButton buttonCopyText;
        internal RibbonButton buttonDeveloper;
        internal RibbonGroup groupDraw;
        internal RibbonButton buttonGraph;
        internal RibbonGroup groupExport;
        internal RibbonButton buttonExportSelection;
        internal RibbonButton buttonSelectionXHTML;
        internal RibbonGroup groupDev;
        internal RibbonButton buttonScrambleText;
        internal RibbonButton buttomResetPageOrigin;
        internal RibbonGroup groupPage;
        internal RibbonButton buttonResizePageToFit;
    }

    partial class ThisRibbonCollection
    {
        internal VPTRibbon VPTRibbon
        {
            get { return this.GetRibbon<VPTRibbon>(); }
        }
    }
}
