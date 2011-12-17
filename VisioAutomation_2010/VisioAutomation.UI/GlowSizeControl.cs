using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class GlowSizeControl : UserControl
    {
        public int GlowSize
        {
            get { return (int) this.numericUpDown1.Value; }
            set
            {
                this.numericUpDown1.Value = value;
                this.sliderGlowSize.Value = value;
            }
        }

        public GlowSizeControl()
        {
            InitializeComponent();
        }


        private void numericUpDown1_ValueChanged(object sender, System.EventArgs e)
        {
            int n = (int) this.numericUpDown1.Value;
            this.sliderGlowSize.Value = n;
        }

        private void ucSliderGlowSize_ValueChanged(object sender, System.EventArgs e)
        {
            int n = (int) this.sliderGlowSize.Value;
            this.numericUpDown1.Value = n;
        }
    }
}