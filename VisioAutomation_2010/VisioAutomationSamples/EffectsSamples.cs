using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using GRIDMODEL = VisioAutomation.Models.Grid;

namespace VisioAutomationSamples
{
    public static class EffectsSamples
    {
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

            SampleEnvironment.SetPageSize(page,page_size);

            var layout = new GRIDMODEL.GridLayout(num_cols, num_rows, new VA.Drawing.Size(6.0, 1.0), master);
            layout.RowDirection = GRIDMODEL.RowDirection.TopToBottom;
            layout.Origin = upperleft;
            layout.CellSpacing = new VA.Drawing.Size(0.1, 0.1);
            layout.PerformLayout();

            double[] trans = EffectsSamples.RangeSteps(0.0, 1.0, num_rows).ToArray();

            int i = 0;
            foreach (var node in layout.Nodes)
            {
                double transparency = trans[i];

                var fmt = new VA.DOM.ShapeCells();
                node.Cells = fmt;

                fmt.FillPattern = 25; // Linear pattern left to right
                fmt.FillForegnd = color1.ToFormula();
                fmt.FillBkgnd = color2.ToFormula();
                fmt.FillForegndTrans = 0;
                fmt.FillBkgndTrans = transparency;
                fmt.LinePattern = 0;

                node.Text = $"bg trans = {transparency}%";
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