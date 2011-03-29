using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools
{
    public partial class FormLayoutTool 
    {
        public FormLayoutTool()
        {
            InitializeComponent();
        }

        private void buttonResizePageToFitContents_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdResizeFitToContents();
        }

        private void buttonDuplicatePage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.Page.DuplicatePage();
        }
    }
}