namespace VisioPowerTools
{
    public partial class FormTextTool 
    {
        public FormTextTool()
        {
            InitializeComponent();
        }

        private void buttonSwitchTextCase_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.CmdSwitchCase();
        }

        private void buttonTextToBottom_Click(object sender, System.EventArgs e)
        {
            // TODO: fix

            VisioPowerToolsAddIn.Client.Text.MoveTextToBottom(null);
        }

        private void buttonResizeToFitText_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Text.FitShapeToText();
        }

        private void buttonEnableTextWrapping_Click(object sender, System.EventArgs e)
        {

            VisioPowerToolsAddIn.Client.Text.SetTextWrapping(true);
        }

        private void buttonDisableTextWrapping_Click(object sender, System.EventArgs e)
        {
            VisioPowerToolsAddIn.Client.Text.SetTextWrapping(false);
        }
    }
}