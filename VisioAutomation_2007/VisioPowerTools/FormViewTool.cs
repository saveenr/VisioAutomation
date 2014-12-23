using VisioAutomation.Scripting;
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
            VisioPowerToolsAddIn.Client.Page.GoTo(PageDirection.Previous);
        }
        
        private void buttonNextPage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Page.GoTo(PageDirection.Next);
        }

        private void buttonZoomToSelection_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdZoomOnSelection();
        }

        private void buttonZoomToPage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.View.Zoom(VA.Scripting.Zoom.ToPage);
        }

        private void buttonZoomIn_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.View.Zoom(VA.Scripting.Zoom.In);
        }

        private void buttonZoomOut_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.View.Zoom(VA.Scripting.Zoom.Out);
        }

        private void buttonFirstPage_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Page.GoTo(PageDirection.First);
        }

        private void buttonPageLast_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Page.GoTo(PageDirection.Last);
        }
    }
}