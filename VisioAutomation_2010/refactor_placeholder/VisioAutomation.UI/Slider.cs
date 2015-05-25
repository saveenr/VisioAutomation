using System.ComponentModel;
using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class Slider : UserControl
    {
        private System.Drawing.Rectangle slider_rect;
        private readonly System.Drawing.Rectangle slider_rect2;
        private int marker_radius = 4;
        private float _min = 0.0f;
        private float _max = 100.0f;
        private readonly System.Drawing.Point[] points;
        private readonly System.Drawing.Point[] groove_points;
        private readonly int groove_y;

        public Slider()
        {
            this.InitializeComponent();
            this.points = new System.Drawing.Point[4];
            int horizontal_padding = 4;
            this.slider_rect = new System.Drawing.Rectangle(horizontal_padding, 0, this.Width - (2 * horizontal_padding) - 1, this.Height - 1);
            this.slider_rect2 = new System.Drawing.Rectangle(0, 0, this.Width - 1, this.Height - 1);
            this.groove_y = this.slider_rect.Top + (int) (this.slider_rect.Height/2.0);
            this.groove_points = new[]
                                {
                                    new System.Drawing.Point(this.slider_rect.Left, this.groove_y),
                                    new System.Drawing.Point(this.slider_rect.Right, this.groove_y)
                                };
        }


        [Browsable(true)]
        public float Min
        {
            get { return this._min; }
            set
            {
                if (value >= this.Max)
                {
                    throw new System.ArgumentOutOfRangeException("must be less than Max value");
                }
                this._min = value;
            }
        }

        [Browsable(true)]
        public float Max
        {
            get { return this._max; }
            set
            {
                if (value <= this.Min)
                {
                    throw new System.ArgumentOutOfRangeException("must be more than Min value");
                }
                this._max = value;
            }
        }


        private float _value;


        [Browsable(true)]
        public float Value
        {
            get { return this._value; }
            set
            {
                this._value = value;
                this.Invalidate();
            }
        }

        private void UCSlider_Paint(object sender, PaintEventArgs e)
        {
            //base.OnPaint(e);
            var rect_pen = System.Drawing.SystemPens.ButtonShadow;
            var groove_pen = System.Drawing.SystemPens.ControlDarkDark;
            var value_pen = System.Drawing.SystemPens.ButtonShadow;
            var value_brush = System.Drawing.SystemBrushes.ButtonShadow;
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

            //e.Graphics.DrawRectangle(rect_pen, slider_rect);
            e.Graphics.DrawRectangle(rect_pen, this.slider_rect2);
            e.Graphics.DrawLine(groove_pen, this.groove_points[0], this.groove_points[1]);

            float scale = this.slider_rect.Width/(this.Max - this.Min);

            int cx = this.slider_rect.Left + (int) (this.Value*scale);
            int cy = (int) (this.slider_rect.Height/2.0) + 1;
            this.points[0] = new System.Drawing.Point(cx, cy);
            this.points[1] = new System.Drawing.Point(cx + this.marker_radius, cy + this.marker_radius);
            this.points[2] = new System.Drawing.Point(cx - this.marker_radius, cy + this.marker_radius);
            this.points[3] = new System.Drawing.Point(cx, cy);

            e.Graphics.FillPolygon(value_brush, this.points);
        }

        private void handle_mouse(int X)
        {
            int new_point = X - this.slider_rect.Left;
            float scale = this.slider_rect.Width/(this.Max - this.Min);

            float new_value = new_point/scale;
            new_value = System.Math.Max(this.Min, new_value);
            new_value = System.Math.Min(new_value, this.Max);

            if (this.Value != new_value)
            {
                this.Value = new_value;
                if (this.ValueChanged != null)
                {
                    var ev_args = new System.EventArgs();
                    this.ValueChanged(this, ev_args);
                }
            }
        }

        private void UCSlider_MouseDown(object sender, MouseEventArgs e)
        {
            this.handle_mouse(e.X);
        }

        public delegate void ValueChangedEventHandler(object sender, System.EventArgs e);

        [Browsable(true)]
        public event ValueChangedEventHandler ValueChanged;

        private void UCSlider_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.handle_mouse(e.X);
            }
        }
    }
}