using System.Linq;

namespace VisioAutomation.UI
{
    internal static class WinFormUtil
    {
        public static void SetItemsChecked(System.Windows.Forms.CheckedListBox checkedlistbox, bool check)
        {
            if (checkedlistbox == null)
            {
                throw new System.ArgumentNullException("checkedlistbox");
            }

            foreach (int index in Enumerable.Range(0, checkedlistbox.Items.Count))
            {
                checkedlistbox.SetItemChecked(index, check);
            }

        }
        public static System.Drawing.Bitmap create_hue_bitmap(int width, int height)
        {
            var bitmap = new System.Drawing.Bitmap(width, height);

            using (var gfx = System.Drawing.Graphics.FromImage(bitmap))
            {
                var colorblend = new System.Drawing.Drawing2D.ColorBlend();
                const int num_steps = 34;
                var range_steps = EnumerableUtil.RangeSteps(0.0, 1.0, num_steps);

                colorblend.Colors = new System.Drawing.Color[num_steps];
                colorblend.Positions = new float[num_steps];

                double _sat = 1.0;
                double _val = 1.0;

                var colors = range_steps.Select(x => ColorUtil.HSVToSystemDrawingColor(x, _sat, _val));
                var positions = range_steps.Select(x => (float) x);

                EnumerableUtil.FillArray( colorblend.Colors, colors );
                EnumerableUtil.FillArray(colorblend.Positions, positions);

                using (var brush_rainbow = new System.Drawing.Drawing2D.LinearGradientBrush(
                    new System.Drawing.Point(0, 0), 
                    new System.Drawing.Point(bitmap.Width, 0),
                    System.Drawing.Color.Black,
                    System.Drawing.Color.White))
                {
                    brush_rainbow.InterpolationColors = colorblend;
                    gfx.FillRectangle(brush_rainbow, 0, 0, bitmap.Width, bitmap.Height);
                }
            }
            return bitmap;
        }

        public static System.Drawing.Bitmap create_hue_bitmap2(int width, int height)
        {
            var bitmap = new System.Drawing.Bitmap(width, height);

            using (var gfx = System.Drawing.Graphics.FromImage(bitmap))
            {
                for (int x = 0; x < width; x++)
                {
                    var h = x / (double)bitmap.Width;
                    double _sat = 1.0;
                    double _val = 1.0;
                    var c0 = ColorUtil.HSVToSystemDrawingColor(h, _sat, _val);
                    uint rgb = (uint) (c0.R << 16 | c0.G << 8 | c0.B);
                    uint mask = 0xff000000;
                    var c2 = System.Drawing.Color.FromArgb((int)(mask | rgb));
                    using (var p = new System.Drawing.Pen(c2))
                    {
                        gfx.DrawLine(p, x, 0, x, bitmap.Height);
                    }
                }
            }
            return bitmap;
        }

    }
}