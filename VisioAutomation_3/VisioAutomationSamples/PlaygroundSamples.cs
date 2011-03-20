using VisioAutomation;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;

namespace VisioAutomationSamples
{
    public static partial class PlaygroundSamples
    {
        private static IVisio.Shape draw_leaf(IVisio.Page page, VA.Drawing.Point p0)
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

        public static VA.Drawing.Point GetPointAtRadius(VA.Drawing.Point origin, double angle, double radius)
        {
            var new_point = new VA.Drawing.Point(radius*System.Math.Cos(angle),
                                         radius*System.Math.Sin(angle));
            new_point = origin + new_point;
            return new_point;
        }

        public static void Spirograph()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            page.Name = "Spirograph";

            var colors =
                new[]
                    {
                        0xf26420, 0xf7931c, 0xfec20d, 0xfff200, 0xcada28, 0x8cc63e, 0x6c9d30,
                        0x288f39,
                        0x6f3a, 0x6f71,
                        0x8eb0, 0xadee, 0x8ed3, 0x71bb, 0x53a6, 0x2e3091, 0x5b57a6, 0x652d91,
                        0x92278e,
                        0xbd198c, 0xec008b,
                        0xec1c23, 0xc1272c, 0x981a1e
                    };

            var origin = new VA.Drawing.Point(4, 4);
            double radius = 3.0;
            int numpoints = colors.Length;
            double angle_step = (System.Math.PI*2/numpoints);
            var angles = Enumerable.Range(0, numpoints).Select(i => i*angle_step).ToList();
            var centers = angles.Select(a => GetPointAtRadius(origin, a, radius)).ToList();
            var shapes = centers.Select(p => draw_leaf(page, p)).ToList();
            var angles_as_formulas = angles.Select(a => a.ToString()).ToList();

            var color_formulas = colors.Select(x => new VA.Drawing.ColorRGB(x).ToFormula()).ToList();

            var shapeids = shapes.Select(s => s.ID16).ToList();

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            var fmt = new VisioAutomation.Format.ShapeFormatCells();
            var xfrm = new VisioAutomation.Layout.XFormCells();

            foreach (int i in Enumerable.Range(0, shapeids.Count))
            {
                short shapeid = shapeids[i];

                xfrm.Angle = angles_as_formulas[i];
                fmt.FillForegnd = color_formulas[i];
                fmt.LineWeight = 0;
                fmt.LinePattern = 0;
                fmt.FillForegndTrans = 0.5;

                xfrm.Apply(update, shapeid);
                fmt.Apply(update, shapeid);
            }

            update.Execute(page);
            VA.PageHelper.ResizeToFitContents(page, new VA.Drawing.Size(1.0, 1.0));
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

            var page_size = new VA.Drawing.Size(5, 5);
            page.SetSize(page_size);

            var lowerleft = new VA.Drawing.Point(0, 0);
            var actual_page_size = page.GetSize();
            var page_rect = new VA.Drawing.Rectangle(lowerleft, actual_page_size);

            var layout = new VA.Layout.Grid.GridLayout(num_cols, num_rows, new VA.Drawing.Size(1, 1), master);
            layout.RowDirection = VA.Layout.Grid.RowDirection.TopToBottom;

            layout.PerformLayout(page_rect.UpperLeft, new VA.Drawing.Size(0, 0));
            int max_grad_id = 40;
            int n = 0;
            foreach (var node in layout.Nodes)
            {
                int grad_id = n%max_grad_id;
                node.Text = grad_id.ToString();
                n++;
            }
            layout.Render(page);

            var color1 = new VA.Drawing.ColorRGB(0xffdddd);
            var color2 = new VA.Drawing.ColorRGB(0x00ffff);

            var src_fillbg = new VA.ShapeSheet.SRC(IVisio.VisSectionIndices.visSectionObject,
                                                   IVisio.VisRowIndices.visRowFill,
                                                   IVisio.VisCellIndices.visFillBkgnd);
            var src_fillfg = new VA.ShapeSheet.SRC(IVisio.VisSectionIndices.visSectionObject,
                                                   IVisio.VisRowIndices.visRowFill,
                                                   IVisio.VisCellIndices.visFillForegnd);
            var src_fillpat = new VA.ShapeSheet.SRC(IVisio.VisSectionIndices.visSectionObject,
                                                    IVisio.VisRowIndices.visRowFill,
                                                    IVisio.VisCellIndices.visFillPattern);

            var src_linepat = new VA.ShapeSheet.SRC(IVisio.VisSectionIndices.visSectionObject,
                                                    IVisio.VisRowIndices.visRowLine,
                                                    IVisio.VisCellIndices.visLinePattern);
            var src_lineweight = new VA.ShapeSheet.SRC(IVisio.VisSectionIndices.visSectionObject,
                                                       IVisio.VisRowIndices.visRowLine,
                                                       IVisio.VisCellIndices.visLineWeight);

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            string color1_formula = color1.ToFormula();
            string color2_formula = color2.ToFormula();

            n = 0;
            foreach (var node in layout.Nodes)
            {
                short shapeid = (short) node.ShapeID;
                int grad_id = n%max_grad_id;

                update.SetFormula(shapeid, src_fillpat, grad_id);
                update.SetFormula(shapeid, src_fillfg, color1_formula);
                update.SetFormula(shapeid, src_fillbg, color2_formula);
                update.SetFormula(shapeid, src_linepat, 0);
                update.SetFormula(shapeid, src_lineweight, 0);
                n++;
            }

            update.Execute(page);
            page.ResizeToFitContents(new VA.Drawing.Size(1, 1));
        }
    }
}