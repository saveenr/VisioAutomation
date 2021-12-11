using System.Globalization;
using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet.Writers;

namespace VSamples
{
    public  class SpirographX : SampleMethodBase
    {
        private static Microsoft.Office.Interop.Visio.Shape draw_leaf(Microsoft.Office.Interop.Visio.Page page, VisioAutomation.Core.Point p0)
        {
            var p1 = p0.Add(1, 1);
            var p2 = p1.Add(1, 0);
            var p3 = p2.Add(1, -1);
            var p4 = p3.Add(-1, -1);
            var p5 = p4.Add(-1, 0);
            var p6 = p5.Add(-1, 1);
            var bezier_points = new[] {p0, p1, p2, p3, p4, p5, p6};

            var s = page.DrawBezier(bezier_points);
            return s;
        }

        public static VisioAutomation.Core.Point GetPointAtRadius(VisioAutomation.Core.Point origin, double angle, double radius)
        {
            var new_point = new VisioAutomation.Core.Point(radius * System.Math.Cos(angle),
                radius * System.Math.Sin(angle));
            new_point = origin + new_point;
            return new_point;
        }

        public override void RunSample()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            page.Name = "Spirograph";

            var colors = new[]
            {
                0xf26420, 0xf7931c, 0xfec20d, 0xfff200,
                0xcada28, 0x8cc63e, 0x6c9d30, 0x288f39,
                0x006f3a, 0x006f71, 0x008eb0, 0x00adee,
                0x008ed3, 0x0071bb, 0x0053a6, 0x2e3091,
                0x5b57a6, 0x652d91, 0x92278e, 0xbd198c,
                0xec008b, 0xec1c23, 0xc1272c, 0x981a1e
            };

            var origin = new VisioAutomation.Core.Point(4, 4);
            double radius = 3.0;
            int numpoints = colors.Length;
            double angle_step = (System.Math.PI * 2 / numpoints);
            var angles = Enumerable.Range(0, numpoints).Select(i => i * angle_step).ToList();
            var centers = angles.Select(a => SpirographX.GetPointAtRadius(origin, a, radius)).ToList();
            var shapes = centers.Select(p => SpirographX.draw_leaf(page, p)).ToList();
            var culture = CultureInfo.InvariantCulture;
            var angles_as_formulas = angles.Select(a => a.ToString(culture)).ToList();

            var color_formulas = colors.Select(x => new VisioAutomation.Models.Color.ColorRgb(x).ToFormula()).ToList();

            var shapeids = shapes.Select(s => s.ID16).ToList();

            var writer = new SidSrcWriter();
            var format = new VisioAutomation.Shapes.ShapeFormatCells();
            var xfrm = new VisioAutomation.Shapes.ShapeXFormCells();

            foreach (int i in Enumerable.Range(0, shapeids.Count))
            {
                short shapeid = shapeids[i];

                xfrm.Angle = angles_as_formulas[i];
                format.FillForeground = color_formulas[i];
                format.LineWeight = 0;
                format.LinePattern = 0;
                format.FillForegroundTransparency = 0.5;

                writer.SetValues(shapeid, xfrm);
                writer.SetValues(shapeid, format);
            }

            writer.Commit(page, VisioAutomation.Core.CellValueType.Formula);

            page.ResizeToFitContents(new VisioAutomation.Core.Size(1.0, 1.0));
        }
    }
}