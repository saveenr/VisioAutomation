using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VisioAutomation.Drawing;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public class DataPoint
    {
        // Relating to the value
        public Double Value;

        // Relating to the Label
        public string Label;

        public DataPoint(double v, string label)
        {
            this.Value = v;
            this.Label = label;
        }
    }

    public class SingleValuePieChartGrid
    {
        public IList<DataPoint> DataPoints;
        public VA.Drawing.ColorRGB CellRectColor = new ColorRGB(0xe0e0e0);
        public VA.Drawing.ColorRGB LineColor = new ColorRGB(0xc0c0c0);

        public void Draw(IVisio.Page page)
        {
            var datapoints = this.DataPoints;

            int rows = 2;
            int cols = 3;
            // ensure rows and colums >= 1

            int max_items = rows*cols;
            // ensure max_items >= datapoints

            var pagesize = VA.PageHelper.GetSize(page);
            var upperleft = new VA.Drawing.Point(0, pagesize.Height);

            var cellsize = new VA.Drawing.Size(2.0, 1.5);

            var tb_fmt = new VA.Text.TextBlockFormatCells();
            tb_fmt.VerticalAlign = 0;
            var origin = upperleft;

            var cellfmt = new VA.Format.ShapeFormatCells();
            cellfmt.FillForegnd = this.CellRectColor.ToFormula();
            cellfmt.LinePattern = 0;
            cellfmt.LineWeight= 0.0;

            var valfmt = new VA.Format.ShapeFormatCells();
            valfmt.FillForegnd = this.LineColor.ToFormula();
            valfmt.LinePattern = 1;
            valfmt.LineWeight = VA.Convert.PointsToInches(1.0);
            valfmt.LineColor = this.LineColor.ToFormula();

            var nonvalfmt = new VA.Format.ShapeFormatCells();
            //nonvalfmt.FillForegnd = this.LineColor.ToFormula();
            nonvalfmt.LinePattern = 1;
            nonvalfmt.LineWeight = VA.Convert.PointsToInches(1.0);
            nonvalfmt.LineColor = this.LineColor.ToFormula();

            var rect_shapes = new List<IVisio.Shape>(datapoints.Count());


            var value_shapes = new List<IVisio.Shape>(datapoints.Count());
            var nonvalue_shapes = new List<IVisio.Shape>(datapoints.Count());

            foreach (int row in Enumerable.Range(0,rows))
            {
                foreach (int col in Enumerable.Range(0,cols))
                {
                    int dp_index = (row*cols) + col;
                    if (dp_index<datapoints.Count())
                    {
                        // Get datapoint
                        var dp = datapoints[dp_index];

                        // Handle background cell
                        var ul = origin.Add(col*cellsize.Width, -row*cellsize.Height);
                        var ll = ul.Add(0, -cellsize.Height);
                        var ur = ll.Add(cellsize.Width, cellsize.Height);
                        var cellrect = new VA.Drawing.Rectangle(ll, ur);
                        var cellshape = page.DrawRectangle(cellrect);
                        cellshape.Text = dp.Label;
                        rect_shapes.Add(cellshape);

                        // draw background
                        var piecenter = cellrect.Center;
                        var pieradius = System.Math.Min(cellrect.Width, cellrect.Height)/4.0;
                        var piedata = new[] {dp.Value, 100.0 - dp.Value};
                        var shapes = VA.Layout.LayoutHelper.DrawPieSlices(page, piecenter, pieradius, piedata);
                        var value_shape = shapes[0];
                        var nonvalue_shape = shapes[1];

                        value_shapes.Add(value_shape);
                        nonvalue_shapes.Add(nonvalue_shape);

                    }
                }
            }

            var update = new VA.ShapeSheet.Update.SIDSRCUpdate();

            // format cells rects
            foreach (var shape in rect_shapes)
            {
                short shapeid = shape.ID16;
                tb_fmt.Apply(update,shapeid);
                cellfmt.Apply(update,shapeid);
            }

            foreach (var shape in value_shapes)
            {
                short shapeid = shape.ID16;
                valfmt.Apply(update, shapeid);               
            }

            foreach (var shape in nonvalue_shapes)
            {
                short shapeid = shape.ID16;
                nonvalfmt.Apply(update, shapeid);
            }

            update.Execute(page);

        }
    }
}
