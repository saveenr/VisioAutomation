using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace VisioAutomation.UI.CommonControls
{
    public partial class ColorSelectorMedium : UserControl
    {
        public ColorSelectorMedium()
        {
            InitializeComponent();
            this.colorSelectorSmall1.ColorChanged += colorSelectorSmall1_ColorChanged;
        }

        void colorSelectorSmall1_ColorChanged(object sender, Color c)
        {
            if (this.ColorChanged !=null)
            {
                this.ColorChanged(sender, c);
            }
        }

        [Browsable(true)]
        public System.Drawing.Color Color
        {
            get { return this.smallColorPicker1.Color; }
            set { this.smallColorPicker1.Color = value; }
        }

        public delegate void ColorChangedEventHandler(object sender, Color c);

        [Browsable(true)]
        public event ColorChangedEventHandler ColorChanged;


    }
}