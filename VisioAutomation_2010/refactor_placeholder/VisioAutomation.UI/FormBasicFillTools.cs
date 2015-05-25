using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class FormBasicFillTools : Form
    {
        public FormBasicFillTools()
        {
            this.InitializeComponent();
        }

        public System.Drawing.Color ForegroundColor
        {
            get
            {
                return this.colorSelectorSmallForeground.Color;
            }
            set
            {
                this.colorSelectorSmallForeground.Color = value;
            }
        }

        public System.Drawing.Color BackgroundColor
        {
            get
            {
                return this.colorSelectorSmallBackground.Color;
            }
            set
            {
                this.colorSelectorSmallBackground.Color = value;
            }
        }


        private void buttonCancel_Click(object sender, System.EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonSwapColors_Click(object sender, System.EventArgs e)
        {
            var temp = this.ForegroundColor;
            this.ForegroundColor = this.BackgroundColor;
            this.BackgroundColor = temp;

        }

        private void buttonOK_Click(object sender, System.EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCopyBgToFg_Click(object sender, System.EventArgs e)
        {
            this.ForegroundColor = this.BackgroundColor;
        }

        private void buttonCopyFgtoBg_Click(object sender, System.EventArgs e)
        {
            this.BackgroundColor = this.ForegroundColor;

        }

        private void FormBasicFillTools_Load(object sender, System.EventArgs e)
        {

        }
    }
}
