using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools
{
    public partial class FormViewTool 
    {
        public FormViewTool()
        {
            InitializeComponent();
        }

        private void buttonPreviousPage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.Page.GoTo(VA.Pages.PageNavigation.PreviousPage);
        }
        
        private void buttonNextPage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.Page.GoTo(VA.Pages.PageNavigation.NextPage);
        }

        private void buttonZoomToSelection_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdZoomOnSelection();
        }

        private void buttonZoomToPage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.View.Zoom(VA.Scripting.Zoom.ToPage);
        }

        private void buttonZoomIn_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.View.Zoom(VA.Scripting.Zoom.In);
        }

        private void buttonZoomOut_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.View.Zoom(VA.Scripting.Zoom.Out);
        }

        private void buttonFirstPage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.Page.GoTo(VA.Pages.PageNavigation.FirstPage);
        }

        private void buttonPageLast_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.ScriptingSession.Page.GoTo(VA.Pages.PageNavigation.LastPage);
        }
    }
}