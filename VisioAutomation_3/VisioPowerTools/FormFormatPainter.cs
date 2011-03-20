using System.Windows.Forms;
using VA=VisioAutomation;
using VAS = VisioAutomation.Scripting;

namespace VisioPowerTools
{
    public partial class FormFormatPainter : Form
    {
        private VisioAutomation.Scripting.FormatPainter fp;
        public FormFormatPainter()
        {
            InitializeComponent();

            this.fp = new VAS.FormatPainter();
        }

        private void buttonCopyFill_Click(object sender, System.EventArgs e)
        {
            this.fp.Copy(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Fill);
        }

        private void buttonPasteFill_Click(object sender, System.EventArgs e)
        {
            this.fp.Paste(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Fill);
        }

        private void buttonCopyLine_Click(object sender, System.EventArgs e)
        {
            this.fp.Copy(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Line);
        }

        private void buttonPasteLine_Click(object sender, System.EventArgs e)
        {
            this.fp.Paste(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Line);
        }

        private void buttonCopyShadow_Click(object sender, System.EventArgs e)
        {
            this.fp.Copy(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Shadow);
        }

        private void buttonPasteShadow_Click(object sender, System.EventArgs e)
        {
            this.fp.Paste(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Shadow);
        }

        private void buttonCopyText_Click(object sender, System.EventArgs e)
        {
            this.fp.Copy(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Character);
        }

        private void buttonPasteText_Click(object sender, System.EventArgs e)
        {
            this.fp.Paste(VisioPowerToolsAddIn.ScriptingSession, VA.Format.FormatCategory.Character);
        }

        private void buttonClear_Click(object sender, System.EventArgs e)
        {
            this.fp.Clear();
        }

        private void buttonCopyAll_Click(object sender, System.EventArgs e)
        {
            this.fp.Copy(VisioPowerToolsAddIn.ScriptingSession);
        }

        private void buttonPasteAll_Click(object sender, System.EventArgs e)
        {
            this.fp.Paste(VisioPowerToolsAddIn.ScriptingSession);
        }
    }
}
