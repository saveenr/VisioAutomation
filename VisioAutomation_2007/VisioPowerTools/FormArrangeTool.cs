using System.Windows.Forms;
using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioPowerTools
{
    public partial class FormArrangeTool : Form
    {
        public FormArrangeTool()
        {
            InitializeComponent();
        }

        private void buttonAlignLeft_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdAlignLeft();
        }

        private void buttonCenter_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdAlignCenter();
        }

        private void buttonRight_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdAlignRight();
        }

        private void buttonTop_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdAlignTop();
        }

        private void buttonMiddle_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdAlignMiddle();
        }

        private void buttonButton_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdAlignBottom();
        }

        private void buttonCopySize_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdCopySize();
        }

        private void buttonPasteSize_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdPasteSize();
        }

        private void buttonPasteWidth_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdPasteWidth();
        }

        private void buttonPasteHeight_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdPasteHeight();
        }

        private void FormArrangeTool_Load(object sender, System.EventArgs e)
        {
        }

        private void buttonDistributeHspacing_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdDistributeHorizontalSpacing();
        }

        private void buttonDistributeHCenter_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdDistributeHorizontalCenter();
        }

        private void buttonDistributeVSpacing_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdDistributeVerticalSpacing();
        }

        private void buttonDistributeVCenter_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdDistributeVerticalMiddle();
        }
    }
}
