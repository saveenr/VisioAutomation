using System.ComponentModel;
using System.Windows.Forms;

namespace VisioAutomation.UI.CommonControls
{
    public partial class ColorSelectorSmall : UserControl
    {
        private ColorSelectorLarge colorform;

        public ColorSelectorSmall()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            colorform = new ColorSelectorLarge();
            colorform.Color = this.Color;
            var popup = new PascalGanaye.Popup.Popup(colorform, this);
            popup.AnimationSpeed = 0;
            popup.DropDownClosed += new System.EventHandler(popup_DropDownClosed);
            popup.Show();
        }

        private void popup_DropDownClosed(object sender, System.EventArgs e)
        {
            if (colorform.ColorSelected)
            {
                this.Color = colorform.Color;
                if (this.ColorChanged !=null)
                {
                    this.ColorChanged(sender, this.Color);
                }
            }
        }


        [Browsable(true)]
        public System.Drawing.Color Color
        {
            get { return this.panelColor.BackColor; }
            set { this.panelColor.BackColor = value; }
        }


        public delegate void ColorChangedEventHandler(object sender, System.Drawing.Color c);

        [Browsable(true)]
        public event ColorChangedEventHandler ColorChanged;

    }
}