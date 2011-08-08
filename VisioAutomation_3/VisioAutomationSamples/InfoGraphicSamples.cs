using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;
using GRID = VisioAutomation.Layout.Grid;
using IG = VisioAutomation.Infographics;

namespace VisioAutomationSamples
{
    public static class InfoGraphicSamples
    {

        public static void DOC()
        {
            var ig = new IG.InfographicDocument();

            var header1 = new IG.Header("JS");
            ig.Blocks.Add(header1);

            var header2 = new IG.Header("JS2");
            ig.Blocks.Add(header2);

            var data = new[] {0.0, 0.25, 0.3, 0.80, 1.0};
            var datapoints = data.Select(i => new IG.DataPoint(i, i.ToString())).ToList();
            var g1 = new IG.SingleValuePieChartGrid();
            g1.DataPoints  = datapoints;
            ig.Blocks.Add(g1);

            var app = SampleEnvironment.Application;
            var doc = app.ActiveDocument;
            ig.RenderPage(doc);
        }

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

            page.ResizeToFitContents(1,1);
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
            page.ResizeToFitContents(1,1);
        }

        public static void PercentGrid()
        {
            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var highlighted_cells = new VA.DOM.ShapeCells();
            highlighted_cells.FillForegnd = "rgb(255,0,0)";
            highlighted_cells.LinePattern = 0;
            highlighted_cells.LineWeight = 0;

            var dimmed_cells = new VA.DOM.ShapeCells();
            dimmed_cells.FillForegnd = "rgb(240,240,240)";
            dimmed_cells.LinePattern = 0;
            dimmed_cells.LineWeight = 0;

            var visapp = page.Application;
            var docs = visapp.Documents;
            var stencil = docs.OpenStencil("basic_u.vss");
            var rectmaster = stencil.Masters["Rectangle"];

            var cellsize = new VA.Drawing.Size(0.5, 0.5);
            var layout = new GRID.GridLayout(10, 10, cellsize, rectmaster);
            layout.CellSpacing = new VA.Drawing.Size(0.25, 0.25);

            layout.PerformLayout();

            int num_highlighted = 2;
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    var node = layout.GetNode(col, row);
                    int i = (row * 10) + col;
                    var fmt = i < num_highlighted ? highlighted_cells : dimmed_cells;
                    node.ShapeCells = fmt;
                }
            }

            layout.Render(page);
            page.ResizeToFitContents(0.5, 0.5);

        }
    }
}