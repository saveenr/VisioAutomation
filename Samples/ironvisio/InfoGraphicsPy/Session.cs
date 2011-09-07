using System;
using System.Collections;
using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;
using IG=InfoGraphicsPy;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;

namespace InfoGraphicsPy
{
    public class Session
    {
        private IVisio.Application app;
        private IVisio.Document doc;
        private IVisio.Document stencil;
        private IVisio.Master rectmaster;
        
        public Session()
        {
            this.app = new IVisio.ApplicationClass();
            this.NewDocument();
        }

        public void NewDocument()
        {
            var docs = this.Application.Documents;
            this.doc = docs.Add("");
            this.stencil = docs.OpenStencil("basic_u.vss");
            var masters = stencil.Masters;
            this.rectmaster = masters["Rectangle"];
        }

        public void NewDocument(double w, double h)
        {
            var docs = this.Application.Documents;
            this.doc = docs.Add("");
        }

        public void NewPage()
        {
            var doc = this.doc;
            doc.Pages.Add();
        }

        public void ResizePageToFit()
        {
            var page = this.Page;
            page.ResizeToFitContents();
        }

        public void ResizePageToFit(double w, double h)
        {
            var page = this.Page;
            page.ResizeToFitContents(new VA.Drawing.Size(w,h));
        }

        public IVisio.Application Application
        {
            get { return this.app; }
        }

        public void TestDraw()
        {
            this.TestDrawPieSlices();
        }

        public void TestDrawPieSlices()
        {
            var page = this.Page;

            double cellwidth = 0.5;
            double hsep = 0.10;
            double vsep = 0.10;
            double cellheight = cellwidth;
            double catheight = 0.5;
            var cats = new[] { "A", "B", "C", "D", "E" };
            var datapoints = new DataPoints(new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 });
            var normalized_values = datapoints.GetNormalizedValues();
            var widths = ConstructPositions(datapoints.Count(), cellwidth, hsep);
            var heights = ConstructPositions(new[] { catheight, cellheight }, vsep);
            var grid = new GridLayout(widths, heights);

            int catrow = 0;
            int barrow = 2;

            var top_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();

            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveAllShapeObjects = true;
            var circle_shapes = new List<VA.DOM.Oval>();
            var slice_shapes = new List<VA.DOM.PieSlice>();
            for (int i = 0; i < datapoints.Count; i++)
            {
                var dp = datapoints[i];
                double start = 0;
                double end = 360*normalized_values[i];
                double radius = top_rects[i].Width/2.0;

                var circle_shape = dom.DrawOval(top_rects[i]);
                circle_shapes.Add(circle_shape);

                var dom_shape = dom.DrawPieSlice(top_rects[i].Center, radius, start, end);
                slice_shapes.Add(dom_shape);
            }
            var cat_shapes = this.DrawRects(dom, cat_rects);

            for (int i = 0; i < datapoints.Count; i++)
            {
                slice_shapes[i].Text = datapoints[i].Text.ToString();
                cat_shapes[i].Text = cats[i];
            }

            foreach (var shape in circle_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = "rgb(255,255,255)";
                cells.LineColor = "rgb(220,220,220)";

            }

            foreach (var shape in slice_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = "rgb(240,240,240)";
                cells.LineColor = "rgb(220,220,220)";

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = "0";
                cells.LineWeight = "0.0";
                cells.LinePattern = "0";
            }
            dom.Render(page);


        }

        public void TestDraw1()
        {
            var page = this.Page;

            double cellwidth = 0.5;
            double hsep = 0.10;
            double vsep = 0.10;
            double cellheight = 4;
            double catheight = 0.5;
            var cats = new[] {"A", "B", "C", "D", "E"};
            var datapoints = new DataPoints(new double[] {1.0, 2.0, 3.0, 4.0, 5.0});
            var normalized_values = datapoints.GetNormalizedValues();
            var widths = ConstructPositions(datapoints.Count(), cellwidth, hsep);
            var heights = ConstructPositions(new[] { catheight, cellheight}, vsep);
            var grid = new GridLayout(widths, heights);
            
            int catrow = 0;
            int barrow = 2;

            var top_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();

            var bar_rects = new List<VA.Drawing.Rectangle>(top_rects.Count);
            for (int i=0;i<top_rects.Count;i++)
            {
                var r = top_rects[i];
                var size = new VA.Drawing.Size(r.Width, normalized_values[i]*cellheight);
                var bar_rect = new VA.Drawing.Rectangle(r.LowerLeft, size);
                bar_rects.Add(bar_rect);
            }
            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveAllShapeObjects = true;
            var bar_shapes = this.DrawRects(dom,bar_rects);
            var cat_shapes = this.DrawRects(dom,cat_rects);

            for (int i = 0; i < datapoints.Count; i++)
            {
                bar_shapes[i].Text = datapoints[i].Text.ToString();
                cat_shapes[i].Text = cats[i];
            }

            foreach (var shape in bar_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = "rgb(240,240,240)";
                cells.LineColor = "rgb(220,220,220)";

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = "0";
                cells.LineWeight = "0.0";
                cells.LinePattern= "0";
            }
            dom.Render(page);


        }

        public List<VA.DOM.Shape> DrawOvals(VA.DOM.Document dom, IList<VA.Drawing.Rectangle> rects)
        {
            var dom_shapes = new List<VA.DOM.Shape>();
            foreach (var rect in rects)
            {
                var dom_shape = dom.DrawOval(rect);
                dom_shape.ShapeCells.Width = rect.Width;
                dom_shape.ShapeCells.Height = rect.Height;
                dom_shapes.Add(dom_shape);
            }

            return dom_shapes;
        }

        public List<VA.DOM.Master> DrawRects(VA.DOM.Document dom, IList<VA.Drawing.Rectangle> rects)
        {
            var dom_shapes = new List<VA.DOM.Master>();
            foreach (var rect in rects)
            {
                var dom_shape = dom.Drop(this.rectmaster, rect.Center);
                dom_shape.ShapeCells.Width = rect.Width;
                dom_shape.ShapeCells.Height = rect.Height;
                dom_shapes.Add(dom_shape);
            }

            return dom_shapes;
        }

        public List<IVisio.Shape> DrawRects(IList<VA.Drawing.Rectangle> rects)
        {
            var dom_shapes = new List<VA.DOM.Master>();
            var dom = new VA.DOM.Document();
            foreach (var rect in rects)
            {
                var dom_shape = dom.Drop(this.rectmaster, rect.Center);
                dom_shape.ShapeCells.Width = rect.Width;
                dom_shape.ShapeCells.Height= rect.Height;
                dom_shapes.Add(dom_shape);
            }

            dom.ResolveAllShapeObjects = true;
            dom.Render(this.Page);

            var shapes = new List<IVisio.Shape>();
            foreach (var dom_shape in dom_shapes)
            {
                shapes.Add(dom_shape.VisioShape);
            }

            return shapes;
        }

        public static List<double> ConstructPositions(int numcols, double width, double sep)
        {
            var iwidths = new List<double>();
            for (int i = 0; i < numcols; i++)
            {
                iwidths.Add(width);
            }
            var widths = ConstructPositions(iwidths, sep);
            return widths;
        }

        public static List<double> ConstructPositions(IList<double> iwidths, double sep)
        {
            int numcols = iwidths.Count;
            var widths = new List<double>();
           
            for (int i = 0; i < numcols; i++)
            {
                widths.Add(iwidths[i]);
                if (i < numcols - 1)
                {
                    widths.Add(sep);
                }
            }
            return widths;
        }

        public IEnumerable<T> SkipOdd<T>(IEnumerable<T> items)
        {
            int i = 0;
            foreach (var item in items)
            {
                if (i % 2 == 1)
                {
                    //
                }
                else
                {
                    yield return item;
                }
                i++;
            }
            
        }


        public IVisio.Page Page
        {
            get { return this.Application.ActivePage; }
        }


        
        
    }
}
