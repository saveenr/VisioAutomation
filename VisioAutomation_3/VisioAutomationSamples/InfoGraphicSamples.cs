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
            var infodoc = new IG.Document();

            var header1 = new IG.Header("Example Infographic Document");
            header1.FontSize = 20.0;
            header1.Bold = true;
            infodoc.Blocks.Add(header1);

            var header2 = new IG.Header("Pie Chart Grid");
            infodoc.Blocks.Add(header2);

            var datapoints = new List<IG.DataPoint>();
            datapoints.Add( new IG.DataPoint(0.0,"alpha"));
            datapoints.Add( new IG.DataPoint(0.25, "beta"));
            datapoints.Add( new IG.DataPoint(0.3, "gamma"));
            datapoints.Add( new IG.DataPoint(0.8, "delta"));
            datapoints.Add( new IG.DataPoint(1.0, "epsilon"));
            var piechartgrid = new IG.PieSliceGrid();

            piechartgrid.DataPoints  = datapoints;
            infodoc.Blocks.Add(piechartgrid);

            var header3 = new IG.Header("Bar Chart");
            infodoc.Blocks.Add(header3);


            var barchart1 = new IG.BarChart();

            barchart1.DataPoints.Add(new IG.DataPoint(100.0, "A"));
            barchart1.DataPoints.Add(new IG.DataPoint(90.0, "B"));
            barchart1.DataPoints.Add(new IG.DataPoint(150.0, "C"));
            barchart1.DataPoints.Add(new IG.DataPoint(130.0, "D"));
            barchart1.DataPoints.Add(new IG.DataPoint(46.0, "E"));

            infodoc.Blocks.Add(barchart1);

            var app = SampleEnvironment.Application;
            var doc = app.ActiveDocument;
            infodoc.RenderPage(doc);
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