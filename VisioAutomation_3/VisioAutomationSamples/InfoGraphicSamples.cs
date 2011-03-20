using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static partial class InfoGraphicSamples
    {
        public static void BarChart()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // get the data and the labels to use
            var data = new[] {1, 1.5, 2.0, 1.7, 1.1};
            var labels = new[] {"a", "b", "c", "d", "e"};

            // draw a rectangle for each value
            // all the rectangles will be drawn on top of each other
            var width = 1.0;
            var rects = data.Select(d => new VA.Drawing.Rectangle(0, 0, width, d));
            var shapes = rects.Select(d => page.DrawRectangle(d)).ToList();

            foreach (int i in Enumerable.Range(0, labels.Count()))
            {
                shapes[i].Text = labels[i];
            }
        }

        public static void PieChart()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // get the data and the labels to use
            var data = new double[] {1, 2, 3, 4, 5, 6};

            var radius = 3.0;
            var center = new VA.Drawing.Point(4, 4);
            VA.Layout.LayoutHelper.DrawPieSlices(page, center, radius, data);
        }
    }
}