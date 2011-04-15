using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class EffectsSamples
    {
        public static void SoftShadow()
        {
            var baserect = new VA.Drawing.Rectangle(1, 1, 5, 5);

            var r = 0.25;

            var rects = new VA.Drawing.Rectangle[3,3];

            rects[0, 0] = new VA.Drawing.Rectangle(baserect.Left - r, baserect.Top, baserect.Left, baserect.Top + r);
            rects[0, 1] = new VA.Drawing.Rectangle(baserect.Left, baserect.Top, baserect.Right, baserect.Top + r);
            rects[0, 2] = new VA.Drawing.Rectangle(baserect.Right, baserect.Top, baserect.Right + r, baserect.Top + r);

            rects[1, 0] = new VA.Drawing.Rectangle(baserect.Left - r, baserect.Bottom, baserect.Left, baserect.Top);
            rects[1, 1] = baserect;
            rects[1, 2] = new VA.Drawing.Rectangle(baserect.Right, baserect.Bottom, baserect.Right + r, baserect.Top);

            rects[2, 0] = new VA.Drawing.Rectangle(baserect.Left - r, baserect.Bottom - r, baserect.Left, baserect.Bottom);
            rects[2, 1] = new VA.Drawing.Rectangle(baserect.Left, baserect.Bottom - r, baserect.Right, baserect.Bottom);
            rects[2, 2] = new VA.Drawing.Rectangle(baserect.Right, baserect.Bottom - r, baserect.Right + r, baserect.Bottom);

            var allrects = new[]
                               {
                                   rects[0, 0], rects[0, 1], rects[0, 2],
                                   rects[1, 0], rects[1, 1], rects[1, 2],
                                   rects[2, 0], rects[2, 1], rects[2, 2],
                               };

            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var model_shapes = allrects.Select(rr => new VA.DOM.Master(master, rr.Center)).ToList();
            var model = new VA.DOM.Document();
            model.Shapes.Add(model_shapes);

            foreach (int i in Enumerable.Range(0, allrects.Length))
            {
                var fmt = new VA.DOM.ShapeCells();
                model_shapes[i].ShapeCells = fmt;

                fmt.Width = allrects[i].Width;
                fmt.Height = allrects[i].Height;
            }

            var shadowfils = new[]
                                 {
                                     39, 30, 38,
                                     27, 1, 25,
                                     37, 28, 36
                                 };

            foreach (int i in Enumerable.Range(0, allrects.Length))
            {
                var fmt = model_shapes[i].ShapeCells;
                fmt.FillPattern = string.Format("guard({0})", shadowfils[i]);
                fmt.FillBkgndTrans = "guard(100%)";
                fmt.FillForegnd = "rgb(0,0,0)";
                fmt.LineWeight = "guard(0)";
                fmt.LinePattern = "guard(0)";
            }

            model.Render(page);

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            update.Execute(page);
            /*
            vi.SelectNone();
            vi.Select(list(shape_ll, shape_bottom, shape_lr, shape_left, shape_middle, shape_right,
                                  shape_ul,
                                  shape_top,
                                  shape_ur));
            //vi.Format.FillCells.Pattern = list(37, 28, 36, 27, 1, 25, 39, 30, 38);
            //vi.Format.LineCells.Pattern = list(0);
            //vi.Format.LineCells.Weight = list(0.0);
            vi.SetFormula("FillBkgndTrans", "guard(100%)");
            vi.SetFormula("FillForegnd", "rgb(0,0,0)");
            vi.SetFormula("ShdwPattern", "guard(0)");
            vi.SetFormula("ShapeShdwType", "guard(0)");


            var g = vi.Group();
            vi.SetFormula("ShdwPattern", "guard(0)");
            vi.SetFormula("ShapeShdwType", "guard(0)");

            var indices = vi.AddControl();
            var index = indices[0];
            vi.SetFormula("Controls.Row_1.X", "Width*.25");
            vi.SetFormula("Controls.Row_1.YCon", "2");

            string corner_w = "GUARD(Sheet.10!Controls.Row_1.X)";

            vi.SelectNone();
            vi.Select(g);
            //vi.SelectSubSelect(list(shape_ll, shape_bottom, shape_lr)); // bottom - row 
            vi.SetFormula("Height", corner_w);
            vi.SetFormula("PinY", "GUARD(Height*0.5)");

            vi.SelectNone();
            vi.Select(g);
            //vi.SelectSubSelect(list(shape_left, shape_middle, shape_right)); // middle - row 
            vi.SetFormula("Height", "GUARD(Sheet.10!Height-(Sheet.10!Controls.Row_1.X*2))");
            vi.SetFormula("Piny", "GUARD(Sheet.10!Height*0.5)");

            vi.SelectNone();
            vi.Select(g);
            //vi.Select.SubSelect(list(shape_ul, shape_top, shape_ur)); // top - row 
            vi.SetFormula("Height", corner_w);
            vi.SetFormula("PinY", "GUARD(Sheet.10!Height-(Sheet.10!Controls.Row_1.X*0.5))");

            vi.SelectNone();
            vi.Select(g);
            //vi.Select.SubSelect(list(shape_ll, shape_left, shape_ul)); // left - col
            vi.SetFormula("Width", corner_w);
            vi.SetFormula("PinX", "GUARD(Width*0.5)");

            vi.SelectNone();
            vi.Select(g);
            //vi.Select.SubSelect(list(shape_bottom, shape_middle, shape_top)); // middle - col
            vi.SetFormula("Width", "GUARD(Sheet.10!Width-(Sheet.10!Controls.Row_1.X*2))");
            vi.SetFormula("PinX", "GUARD(Sheet.10!Width*0.5)");

            vi.SelectNone();
            vi.Select(g);
            //vi.Select.SubSelect(list(shape_lr, shape_right, shape_ur)); // left - col
            vi.SetFormula("Width", corner_w);
            vi.SetFormula("PinX", "GUARD(Sheet.10!Width-(Sheet.10!Controls.Row_1.X*0.5))");

            vi.SelectNone();
            vi.Select(g);
            //vi.SelectMode = list(0);
            */
        }

        public static void GradientTransparencies()
        {
            int num_cols = 1;
            int num_rows = 10;
            var color1 = new VA.Drawing.ColorRGB(0xff000);
            var color2 = new VA.Drawing.ColorRGB(0x000ff);

            var page_size = new VA.Drawing.Size(num_rows/2.0, num_rows);
            var upperleft = new VA.Drawing.Point(0, page_size.Height);

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var app = page.Application;
            var docs = app.Documents;
            var stencil = docs.OpenStencil("basic_U.vss");
            var master = stencil.Masters["Rectangle"];

            page.SetSize(page_size);

            var layout = new VA.Layout.Grid.GridLayout(num_cols, num_rows, new VA.Drawing.Size(6.0, 1.0), master);
            layout.RowDirection = VA.Layout.Grid.RowDirection.TopToBottom;
            layout.Origin = upperleft;
            layout.CellSpacing = new VA.Drawing.Size(0.1, 0.1);
            layout.PerformLayout();

            double[] trans = RangeSteps(0.0, 1.0, num_rows).ToArray();

            int i = 0;
            foreach (var node in layout.Nodes)
            {
                double transparency = trans[i];

                var fmt = new VisioAutomation.DOM.ShapeCells();
                node.ShapeCells = fmt;

                fmt.FillPattern = (int)VA.Format.FillPattern.LinearLeftToRight;
                fmt.FillForegnd = color1.ToFormula();
                fmt.FillBkgnd = color2.ToFormula();
                fmt.FillForegndTrans = 0;
                fmt.FillBkgndTrans = transparency;
                fmt.LinePattern = 0;

                node.Text = string.Format("bg trans = {0}%", transparency);
                i++;
            }

            layout.Render(page);

            page.ResizeToFitContents();
        }

        /// <summary>
        /// Given a range (start,end) and a number of steps, will yield that a number for each step
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <param name="steps"></param>
        /// <returns></returns>
        public static IEnumerable<double> RangeSteps(double start, double end, int steps)
        {
            // for non-positive number of steps, yield no points
            if (steps < 1)
            {
                yield break;
            }

            // for exactly 1 step, yield the start value
            if (steps == 1)
            {
                yield return start;
                yield break;
            }

            // for exactly 2 stesp, yield the start value, and then the end value
            if (steps == 2)
            {
                yield return start;
                yield return end;
                yield break;
            }

            // for 3 steps or above, start yielding the segments
            // notice that the start and end values are explicitly returned so that there
            // is no possibility of rounding error affecting their values
            int segments = steps - 1;
            double total_length = end - start;
            double stepsize = total_length/segments;
            yield return start;
            for (int i = 1; i < (steps - 1); i++)
            {
                double p = start + (stepsize*i);
                yield return p;
            }
            yield return end;
        }
    }
}