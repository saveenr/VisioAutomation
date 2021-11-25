using System.Globalization;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using VisioAutomation.Models.Layouts.Grid;
using VisioAutomation.ShapeSheet.Writers;

namespace VisioAutomationSamples
{
    public static class PlaygroundSamples
    {
        private static IVisio.Shape draw_leaf(IVisio.Page page, VA.Geometry.Point p0)
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

        public static VA.Geometry.Point GetPointAtRadius(VA.Geometry.Point origin, double angle, double radius)
        {
            var new_point = new VA.Geometry.Point(radius*System.Math.Cos(angle),
                                         radius*System.Math.Sin(angle));
            new_point = origin + new_point;
            return new_point;
        }

        public static void Spirograph()
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

            var origin = new VA.Geometry.Point(4, 4);
            double radius = 3.0;
            int numpoints = colors.Length;
            double angle_step = (System.Math.PI*2/numpoints);
            var angles = Enumerable.Range(0, numpoints).Select(i => i*angle_step).ToList();
            var centers = angles.Select(a => PlaygroundSamples.GetPointAtRadius(origin, a, radius)).ToList();
            var shapes = centers.Select(p => PlaygroundSamples.draw_leaf(page, p)).ToList();
            var culture = CultureInfo.InvariantCulture;
            var angles_as_formulas = angles.Select(a => a.ToString(culture)).ToList();

            var color_formulas = colors.Select(x => new VA.Models.Color.ColorRgb(x).ToFormula()).ToList();

            var shapeids = shapes.Select(s => s.ID16).ToList();

            var writer = new SidSrcWriter();
            var format = new VA.Shapes.ShapeFormatCells();
            var xfrm = new VA.Shapes.ShapeXFormCells();

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

            writer.Commit(page, VA.ShapeSheet.CellValueType.Formula);

            page.ResizeToFitContents(new VA.Geometry.Size(1.0, 1.0));
        }

        public static void DrawAllGradients()
        {
            var app = SampleEnvironment.Application;
            var docs = app.Documents;
            var stencil = docs.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            int num_cols = 7;
            int num_rows = 7;

            var page_size = new VA.Geometry.Size(5, 5);
            SampleEnvironment.SetPageSize(page,page_size);

            var lowerleft = new VA.Geometry.Point(0, 0);
            var actual_page_size = SampleEnvironment.GetPageSize(page);
            var page_rect = new VA.Geometry.Rectangle(lowerleft, actual_page_size);

            var layout = new GridLayout(num_cols, num_rows, new VA.Geometry.Size(1, 1), master);
            layout.RowDirection = RowDirection.TopToBottom;
            layout.Origin = page_rect.UpperLeft;
            layout.CellSpacing = new VA.Geometry.Size(0, 0);
            layout.PerformLayout();

            int max_grad_id = 40;
            int n = 0;
            
            foreach (var node in layout.Nodes)
            {
                int grad_id = n%max_grad_id;
                node.Text = grad_id.ToString();
                n++;
            }

            layout.Render(page);

            var color1 = new VA.Models.Color.ColorRgb(0xffdddd);
            var color2 = new VA.Models.Color.ColorRgb(0x00ffff);

            var format = new VA.Shapes.ShapeFormatCells();

            var writer = new SidSrcWriter();

            string color1_formula = color1.ToFormula();
            string color2_formula = color2.ToFormula();

            n = 0;

            foreach (var node in layout.Nodes)
            {
                short shapeid = node.ShapeID;
                int grad_id = n%max_grad_id;

                format.FillPattern = grad_id;
                format.FillForeground = color1_formula;
                format.FillBackground = color2_formula;
                format.LinePattern = 0;
                format.LineWeight = 0;

                writer.SetValues(shapeid, format);

                n++;
            }

            writer.Commit(page, VA.ShapeSheet.CellValueType.Formula);

            var bordersize = new VA.Geometry.Size(1, 1);
            page.ResizeToFitContents(bordersize);
        }
    }
}