using VA=VisioAutomation;

namespace VisioAutomation.UI
{
    using System.ComponentModel;
    using System.Windows.Forms;

    public enum DirectionRelative
    {
        Up,
        Right,
        Left,
        Down
    }

    public partial class ThreePointFillControl : UserControl
    {
        [Browsable(true)]
        public System.Drawing.Color Corner1Color
        {
            get { return this.ColorPickerCorner1.Color; }
            set { this.ColorPickerCorner1.Color = value; }
        }

        [Browsable(true)]
        public System.Drawing.Color Corner2Color
        {
            get { return this.ColorPickerCorner2.Color; }
            set { this.ColorPickerCorner2.Color = value; }
        }

        [Browsable(true)]
        public System.Drawing.Color EdgeColor
        {
            get { return this.ColorPickerPrimaryEdge.Color; }
            set { this.ColorPickerPrimaryEdge.Color = value; }
        }

        [Browsable(true)]
        public DirectionRelative Direction
        {
            get
            {
                if (this.radioButtonUp.Checked)
                {
                    return DirectionRelative.Up;
                }
                else if (this.radioButtonRight.Checked)
                {
                    return DirectionRelative.Right;
                }
                else if (this.radioButtonDown.Checked)
                {
                    return DirectionRelative.Down;
                }
                else if (this.radioButtonLeft.Checked)
                {
                    return DirectionRelative.Left;
                }
                else
                {
                    throw new System.ArgumentOutOfRangeException();
                }
            }

            set
            {
                this.radioButtonUp.Checked = (value == DirectionRelative.Up);
                this.radioButtonDown.Checked = (value == DirectionRelative.Down);
                this.radioButtonLeft.Checked = (value == DirectionRelative.Left);
                this.radioButtonRight.Checked = (value == DirectionRelative.Right);
            }
        }

        public ThreePointFillControl()
        {
            InitializeComponent();

            this.Direction =DirectionRelative.Right;
        }

        private void UC3PointFill_Load(object sender, System.EventArgs e)
        {
        }

        private void groupBoxDirection_Enter(object sender, System.EventArgs e)
        {
        }

        private void buttonSwapCorner_Click(object sender, System.EventArgs e)
        {
            var temp = this.Corner1Color;
            this.Corner1Color = this.Corner2Color;
            this.Corner2Color = temp;
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            var new_edge = this.Corner2Color;
            var new_c2 = this.Corner1Color;
            var new_c1 = this.EdgeColor;

            this.EdgeColor = new_edge;
            this.Corner1Color = new_c1;
            this.Corner2Color = new_c2;
        }
    }
}