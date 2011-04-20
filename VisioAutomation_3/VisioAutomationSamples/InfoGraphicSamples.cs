using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class InfoGraphicSamples
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

            double curx = 0;
            var rects = new List<VA.Drawing.Rectangle>(data.Count());
            for (int i = 0; i < data.Count(); i++)
            {
                var rect = new VA.Drawing.Rectangle(curx,0,curx+width,data[i]);
                rects.Add(rect);
                curx += width + 0.5;
            }
            var shapes = rects.Select(d => page.DrawRectangle(d)).ToList();

            foreach (int i in Enumerable.Range(0, labels.Count()))
            {
                shapes[i].Text = labels[i];
            }

            page.ResizeToFitContents(new VA.Drawing.Size(1,1));
        }

        public static void PieChart()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            // get the data and the labels to use
            var data = new double[] {1, 2, 3, 4, 56};
            var colors = new string[] { "rgb(239,233,195)", "rgb(200,233,167)", "rgb(172,208,180)", "rgb(113,121,118)", "rgb(93,70,51)"};
            var radius = 3.0;
            var center = new VA.Drawing.Point(4, 4);
            var shapes = VA.Layout.LayoutHelper.DrawPieSlices(page, center, radius, data);

            var fmt = new VA.Format.ShapeFormatCells();

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();
            for (int i = 0; i < data.Count(); i++)
            {
                var shape = shapes[i];
                var color = colors[i];
                fmt.FillForegnd = color;
                fmt.LinePattern = 0;
                fmt.LineWeight = 0;
                fmt.Apply(update,shape.ID16);
            }
            update.Execute(page);
            page.ResizeToFitContents(new VA.Drawing.Size(1, 1));
        }
    }
}