using System.ComponentModel;
using System.Windows.Forms;

namespace VisioAutomation.UI
{
    public partial class ColorSelectorLarge : UserControl
    {
        public bool ColorSelected;
        public System.Drawing.Bitmap bmp_hue;
        public System.Drawing.Bitmap bmp_gradient;
        private System.Drawing.Point? hue_selection_point;
        private System.Drawing.Point? gradient_selection_point;
        double? gradient_hue;

        public ColorSelectorLarge()
        {
            this.InitializeComponent();
            this.ColorSelected = false;
            
            // Setup the HUE bitmap
            this.bmp_hue = this.bmp_hue ?? WinFormUtil.create_hue_bitmap2(this.pictureBoxHue.Width, this.pictureBoxHue.Height);
            this.pictureBoxHue.Image = this.bmp_hue;
            
            // Setup the GRADIENT BITMAP
            this.bmp_gradient = new System.Drawing.Bitmap(this.pictureBoxGradient.Width, this.pictureBoxGradient.Height);
            this.pictureBoxGradient.Image = this.bmp_gradient;
        }

        [Browsable(true)]
        public System.Drawing.Color Color
        {
            get { return this.panelColor.BackColor; }
            set { this.panelColor.BackColor = value; }
        }
        
        private void buttonOK_Click(object sender, System.EventArgs e)
        {
            this.ColorSelected = true;
            this.Parent.Hide();
        }

        private void pictureBoxHue_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
            this.handle_hue_click(new System.Drawing.Point (e.X, e.Y));
            }
        }

        private void pictureBoxHue_MouseDown(object sender, MouseEventArgs e)
        {
            this.handle_hue_click(new System.Drawing.Point (e.X, e.Y));
        }

        private int clamp(int val, int min, int max)
        {
            return System.Math.Min(max, System.Math.Max(min, val));
        }

        private System.Drawing.Point clamp(System.Drawing.Point p, System.Drawing.Point min, System.Drawing.Point max)
        {
            var cp = new System.Drawing.Point(this.clamp(p.X, min.X, max.X ), this.clamp(p.Y,min.Y, max.Y ) );
            return cp;
        }

        private void handle_hue_click(System.Drawing.Point hue_click)
        {
            // Calculate point in hue bitmap
            var min_point = new System.Drawing.Point(0, 0);
            var max_point = new System.Drawing.Point(this.bmp_hue.Width - 2, this.bmp_hue.Height - 2);
            this.hue_selection_point = this.clamp(hue_click, min_point, max_point);

            // Retrieve the color
            var color_from_hue_slider = this.bmp_hue.GetPixel(this.hue_selection_point.Value.X, this.hue_selection_point.Value.Y);

            this.Color = color_from_hue_slider;

            double _h;
            double _s;
            double _v;

            ColorUtil.RGBToHSV(this.Color, out _h, out _s, out _v);
            this.gradient_hue = _h;
            this.pictureBoxHue.Invalidate();


            this.SetColorFromGradientPoint();
            this.pictureBoxGradient.Invalidate();
        }
        
        private void buttonClose_Click(object sender, System.EventArgs e)
        {
            this.ColorSelected = false;
            this.Parent.Hide();
        }

        private void pictureBoxHue_Paint(object sender, PaintEventArgs e)
        {
            var gfx = e.Graphics;
            if (!this.hue_selection_point.HasValue)
            {
                // Initial Point for Hue

                double _h;
                double _s;
                double _v;

                ColorUtil.RGBToHSV(this.Color, out _h, out _s, out _v);

                this.hue_selection_point = new System.Drawing.Point((int)(_h * this.bmp_hue.Width), 0);
            }
            float cpx = this.hue_selection_point.Value.X;
            float cpy = ((float) this.bmp_hue.Height-2)/2.0f - 0.5f;
            gfx.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            this.draw_cursor_ring(gfx, cpx, cpy);
        }

        private void draw_cursor_ring(System.Drawing.Graphics gfx, float cpx, float cpy)
        {
            float radius = 7;
            float outer_width = 3.0f;
            float inner_width = 1.5f;
            var outer_color = System.Drawing.Color.Black;
            var inner_color = System.Drawing.Color.White;
            var cursor_rect = new System.Drawing.RectangleF(cpx - radius, cpy - radius, radius * 2, radius * 2);
            using (var inner_pen = new System.Drawing.Pen(inner_color, inner_width))
            using (var outer_pen = new System.Drawing.Pen(outer_color, outer_width))
            {
                gfx.DrawEllipse(outer_pen, cursor_rect);
                gfx.DrawEllipse(inner_pen, cursor_rect);
            }
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }


            if (disposing && this.bmp_hue != null)
            {
                this.bmp_hue.Dispose();
            }

            base.Dispose(disposing);
        }

        private System.Drawing.Rectangle GetGradientRect()
        {
            return new System.Drawing.Rectangle(0, 0, this.pictureBoxGradient.Width - 2, this.pictureBoxGradient.Height - 2);
        }

        private System.Drawing.Drawing2D.LinearGradientBrush GetLuminanceBrush()
        {
            // Setup the colors & dimensions of the brush
            var rect = this.GetGradientRect();
            var top_color = System.Drawing.Color.White;
            var bottom_color = System.Drawing.Color.Black;

            // Create the brush object
            var luminance_brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                rect,
                top_color,
                bottom_color,
                System.Drawing.Drawing2D.LinearGradientMode.Vertical);

            return luminance_brush;
        }

        /// <summary>
        /// This retrieved the transparent to fully staurated color part of the gradient
        /// </summary>
        /// <returns></returns>
        private System.Drawing.Drawing2D.LinearGradientBrush GetHueBrush()
        {
            // Aquire a Hue
            if (!this.gradient_hue.HasValue)
            {
                // set initial  hue

                double _h;
                double _s;
                double _v;

                ColorUtil.RGBToHSV(this.Color, out _h, out _s, out _v);
                this.gradient_hue = _h;
            }

            // Setup the colors & dimensions of the brush
            var _hue = this.gradient_hue.Value;
            var _sat = 1.0;
            var _val = 1.0;
            var left_color = System.Drawing.Color.Transparent;
            var right_color = ColorUtil.HSVToSystemDrawingColor(_hue, _sat, _val);
            var rect = this.GetGradientRect();

            // Create the brush object
            var hue_brush = new System.Drawing.Drawing2D.LinearGradientBrush(
                rect,
                left_color,
                right_color,
                System.Drawing.Drawing2D.LinearGradientMode.Horizontal);

            return hue_brush;
        }

        private void pictureBoxGradient_Paint(object sender, PaintEventArgs e)
        {
            using (var gfx = System.Drawing.Graphics.FromImage(this.bmp_gradient))
            {
                this.Gradient_Paint(gfx);
            }
        }

        private void Gradient_Paint(System.Drawing.Graphics gfx)
        {
            var gradient_rect = this.GetGradientRect();

            using (var luminance_brush = this.GetLuminanceBrush())
            using (var hue_brush = this.GetHueBrush())
            {
                // draw the hsv gradient
                gfx.FillRectangle(luminance_brush, gradient_rect);
                gfx.FillRectangle(hue_brush, gradient_rect);

                if (!this.gradient_selection_point.HasValue)
                {
                    double _h;
                    double _s;
                    double _v;

                    ColorUtil.RGBToHSV(this.Color, out _h, out _s, out _v);
                    
                    int x = (int)(_s * (this.bmp_gradient.Width - 1));
                    int y = (int)( (1.0 - _v) * (this.bmp_gradient.Height - 1));
                    this.gradient_selection_point = new System.Drawing.Point(x, y);
                }

                // draw the cursor
                gfx.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                var p = this.gradient_selection_point.Value;
                this.draw_cursor_ring(gfx, p.X , p.Y);

            }
        }

        private bool is_selecting_gradient = false;
        private void pictureBoxGradient_MouseDown(object sender, MouseEventArgs e)
        {
            this.is_selecting_gradient = true;
            this.handle_gradient_click(ColorSelectorLarge.point(e));
        }

        private void pictureBoxGradient_MouseMove(object sender, MouseEventArgs e)
        {
            if (this.is_selecting_gradient)
            {
                this.handle_gradient_click(ColorSelectorLarge.point(e));
            }
        }

        private void pictureBoxGradient_MouseUp(object sender, MouseEventArgs e)
        {
            this.handle_gradient_click( ColorSelectorLarge.point(e));
            this.is_selecting_gradient = false;

        }

        public void handle_gradient_click(System.Drawing.Point p)
        {
            var min_point = new System.Drawing.Point(0, 0);
            var max_point = new System.Drawing.Point(this.bmp_gradient.Width - 3, this.bmp_gradient.Height - 3);
            this.gradient_selection_point = this.clamp(p, min_point, max_point);

            this.SetColorFromGradientPoint();

            this.pictureBoxGradient.Invalidate();
        }

        private void SetColorFromGradientPoint()
        {
            System.Drawing.Point p = this.gradient_selection_point.Value;
            var c = this.bmp_gradient.GetPixel(p.X, p.Y);
            this.Color = c;
        }

        static System.Drawing.Point point( MouseEventArgs e)
        {
            return new System.Drawing.Point(e.X, e.Y);
        }

        private bool bmp_gradient_init = false;

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (!this.bmp_gradient_init)
            {
                using (var gfx = System.Drawing.Graphics.FromImage(this.bmp_gradient))
                {
                    this.Gradient_Paint(gfx);
                }
                this.bmp_gradient_init = true;
            }

        }


        private void ColorSelectorLarge_Paint(object sender, PaintEventArgs e)
        {
        }
    }
}