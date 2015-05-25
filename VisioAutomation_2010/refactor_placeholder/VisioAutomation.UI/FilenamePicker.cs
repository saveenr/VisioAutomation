using System.ComponentModel;
using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class FilenamePicker : UserControl
    {
        private const int MAX_LINES = 1;

        readonly OpenFileDialog openfiledialog = new OpenFileDialog();

        public FilenamePicker()
        {
            this.InitializeComponent();
        }

        public OpenFileDialog OpenFileDialog
        {
            get
            {
                return this.openfiledialog;
            }
        }

        public string Filename
        {
            get
            {
                return this.Text;
            }
            set
            {
                this.Text = value;
            }
        }

        [Browsable(true)]
        public override string Text
        {
            get
            {
                return this.textBoxFilename.Text.Trim();
            }
            set
            {
                this.textBoxFilename.Text = value;
            }
        }

        [Browsable(true)]
        public bool ReadOnly
        {
            get
            {
                return this.textBoxFilename.ReadOnly;
            }
            set
            {
                this.textBoxFilename.ReadOnly = value;
            }
        }


        private void textBoxFilename_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (this.textBoxFilename.Lines.Length >= FilenamePicker.MAX_LINES && e.KeyChar == '\r')
            {
                e.Handled = true;
            } 

        }

        private void textBoxFilename_TextChanged(object sender, System.EventArgs e)
        {
            if (this.textBoxFilename.Lines.Length > 1)
            {
                string first_line = this.textBoxFilename.Lines[0];
                this.textBoxFilename.Text = first_line.Trim();
            } 
        }

        private void buttonBrowse_Click(object sender, System.EventArgs e)
        {

            var result = this.openfiledialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                this.Filename = this.openfiledialog.FileName;
            }

        }
    }
}
