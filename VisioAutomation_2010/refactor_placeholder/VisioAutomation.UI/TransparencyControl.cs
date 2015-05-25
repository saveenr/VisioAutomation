using System.ComponentModel;
using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class TransparencyControl : UserControl
    {
        [Browsable(true)]
        public int TransparencyPercent
        {
            get { return (int) this.numericUpDown1.Value; }
            set
            {
                this.numericUpDown1.Value = value;
                this.slider1.Value = value;
            }
        }

        public TransparencyControl()
        {
            this.InitializeComponent();
        }


        private void numericUpDown1_ValueChanged(object sender, System.EventArgs e)
        {
            int n = (int) this.numericUpDown1.Value;
            this.slider1.Value = n;
        }


        private void ucSlider1_ValueChanged(object sender, System.EventArgs e)
        {
            int n = (int) this.slider1.Value;
            this.numericUpDown1.Value = n;
        }
    }
}