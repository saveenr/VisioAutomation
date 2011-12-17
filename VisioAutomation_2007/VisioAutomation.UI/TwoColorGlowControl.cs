using System.ComponentModel;
using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class TwoColorGlowControl : UserControl
    {
        [Browsable(true)]
        public System.Drawing.Color UpperColor
        {
            get { return this.colorPickerUpperGlow.Color; }
            set { this.colorPickerUpperGlow.Color = value; }
        }

        [Browsable(true)]
        public System.Drawing.Color LowerColor
        {
            get { return this.colorPickerLowerGlow.Color; }
            set { this.colorPickerLowerGlow.Color = value; }
        }

        [Browsable(true)]
        public int UpperTransparency
        {
            get { return this.transparency1.TransparencyPercent; }
            set { this.transparency1.TransparencyPercent = value; }
        }

        [Browsable(true)]
        public int LowerTransparency
        {
            get { return this.transparency2.TransparencyPercent; }
            set { this.transparency2.TransparencyPercent = value; }
        }

        [Browsable(true)]
        public int GlowSize
        {
            get { return this.glowSize1.GlowSize; }
            set { this.glowSize1.GlowSize = value; }
        }

        public TwoColorGlowControl()
        {
            InitializeComponent();
        }
    }
}