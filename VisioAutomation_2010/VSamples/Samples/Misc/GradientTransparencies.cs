using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Extensions;
using VA = VisioAutomation;
using VAM = VisioAutomation.Models;

namespace VSamples.Samples.Misc
{
    public class GradientTransparencies : SampleMethodBase
    {
        public override void Run()
        {
            int num_cols = 1;
            int num_rows = 10;
            var color1 = new VAM.Color.ColorRgb(0xff000);
            var color2 = new VAM.Color.ColorRgb(0x000ff);

            var page_size = new VA.Core.Size(num_rows / 2.0, num_rows);
            var upperleft = new VA.Core.Point(0, page_size.Height);

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();
            var app = page.Application;
            var docs = app.Documents;
            var stencil = docs.OpenStencil("basic_U.vss");
            var master = stencil.Masters["Rectangle"];

            SampleEnvironment.SetPageSize(page, page_size);

            var layout = new VAM.Layouts.Grid.GridLayout(num_cols, num_rows, new VA.Core.Size(6.0, 1.0), master);
            layout.RowDirection = VAM.Layouts.Grid.RowDirection.TopToBottom;
            layout.Origin = upperleft;
            layout.CellSpacing = new VA.Core.Size(0.1, 0.1);
            layout.PerformLayout();

            double[] trans = GradientTransparencies.RangeSteps(0.0, 1.0, num_rows).ToArray();

            int i = 0;
            foreach (var node in layout.Nodes)
            {
                double transparency = trans[i];

                var fmt = new VAM.Dom.ShapeCells();
                node.Cells = fmt;

                fmt.FillPattern = 25; // Linear pattern left to right
                fmt.FillForeground = color1.ToFormula();
                fmt.FillBackground = color2.ToFormula();
                fmt.FillForegroundTransparency = 0;
                fmt.FillBackgroundTransparency = transparency;
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
            double stepsize = total_length / segments;
            yield return start;
            for (int i = 1; i < (steps - 1); i++)
            {
                double p = start + (stepsize * i);
                yield return p;
            }

            yield return end;
        }
    }
}