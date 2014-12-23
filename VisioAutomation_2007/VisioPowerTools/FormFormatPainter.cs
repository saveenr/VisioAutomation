using System.Windows.Forms;
using VisioAutomation.Scripting;
using VA=VisioAutomation;

namespace VisioPowerTools
{
    public partial class FormFormatPainter : Form
    {
        public FormFormatPainter()
        {
            InitializeComponent();

        }

        private void buttonCopyFill_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Copy(null, FormatCategory.Fill);
        }

        private void buttonPasteFill_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Paste(null, FormatCategory.Fill,false);
        }

        private void buttonCopyLine_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Copy(null, FormatCategory.Line);
        }

        private void buttonPasteLine_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Paste(null, FormatCategory.Line, false);
        }

        private void buttonCopyShadow_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Copy(null, FormatCategory.Shadow);
        }

        private void buttonPasteShadow_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Paste(null, FormatCategory.Shadow, false);
        }

        private void buttonCopyText_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Copy(null, FormatCategory.Character);
        }

        private void buttonPasteText_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Paste(null, FormatCategory.Character, false);
        }

        private void buttonClear_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.ClearFormatCache();
        }

        private void buttonCopyAll_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Copy();
        }

        private void buttonPasteAll_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Format.Paste(null, FormatCategory.Character | FormatCategory.Fill |FormatCategory.Line|FormatCategory.Paragraph| FormatCategory.Shadow,false);
        }
    }
}
