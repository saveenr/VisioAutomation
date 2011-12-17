using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerTools
{
    public partial class FormSelectionTool
    {
        public FormSelectionTool()
        {
            InitializeComponent();
        }

        private void buttonSelectAll_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdSelectAll();
        }

        private void buttonSelectNone_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdSelectNone();
        }

        private void buttonInvertSelection_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdInvertSelection();
        }
    }
}