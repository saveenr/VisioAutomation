using System;
using System.Collections.Generic;
using System.Linq;
using VA=VisioAutomation;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Infographics
{
    public class PieSliceGrid : Block
    {
        public IList<DataPoint> DataPoints;
        public VA.Drawing.ColorRGB ValueColor = new VA.Drawing.ColorRGB(0xa0a0a0);
        public VA.Drawing.ColorRGB NonValueColor = new VA.Drawing.ColorRGB(0xffffff);

        public override VA.Drawing.Size Render(RenderContext rc)
        {
            var page = rc.Page;
            var datapoints = this.DataPoints;

            var grid = new GridBuilder(2, 3);

            if (datapoints.Count>grid.Count)
            {
                throw new System.ArgumentOutOfRangeException("Too many datapoints to fit into grid");
            }

            var margin = new VA.Drawing.Size(0.25, 0.25);

            var grid_size = grid.Size;
            var grid_size_actual = grid_size.Add(margin).Add(margin);
            var max_width = System.Math.Max(grid_size.Width, rc.PageWidth);
            var bkrect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, new VA.Drawing.Size(max_width,grid_size_actual.Height));

            var bkshape = page.DrawRectangle(bkrect);

            var bkfmt = rc.GetDefaultBkfmt();

            var tb_fmt = new VA.Text.TextBlockFormatCells();
            tb_fmt.VerticalAlign = 0;
            var origin = bkrect.UpperLeft.Add(margin.Width, -margin.Height);

            var cellfmt = new VA.Format.ShapeFormatCells();
            cellfmt.FillPattern = 0;
            cellfmt.LinePattern = 0;
            cellfmt.LineWeight = 0.0;

            var cellcharfmt = new VA.Text.CharacterFormatCells();
            cellcharfmt.Font = rc.GetFontID(rc.DefaultFont);

            var valfmt = new VA.Format.ShapeFormatCells();
            valfmt.FillForegnd = this.ValueColor.ToFormula();
            valfmt.LinePattern = 1;
            valfmt.LineWeight = VA.Convert.PointsToInches(1.0);
            valfmt.LineColor = rc.LineColor.ToFormula();

            var nonvalfmt = new VA.Format.ShapeFormatCells();
            nonvalfmt.FillForegnd = this.NonValueColor.ToFormula();
            nonvalfmt.LinePattern = 1;
            nonvalfmt.LineWeight = VA.Convert.PointsToInches(1.0);
            nonvalfmt.LineColor = rc.LineColor.ToFormula();

            var cell_shapes = new List<IVisio.Shape>(datapoints.Count());


            var value_shapes = new List<IVisio.Shape>(datapoints.Count());
            var nonvalue_shapes = new List<IVisio.Shape>(datapoints.Count());

            foreach (int row in Enumerable.Range(0, grid.RowCount))
            {
                foreach (int col in Enumerable.Range(0, grid.ColumnCount))
                {
                    int dp_index = (row * grid.ColumnCount) + col;
                    if (dp_index < datapoints.Count())
                    {
                        // Get datapoint
                        var dp = datapoints[dp_index];

                        // Handle background cell
                        var cellrect = grid.GetCellRect(origin, row,col);
                        var cellshape = page.DrawRectangle(cellrect);
                        cellshape.Text = dp.Label;
                        cell_shapes.Add(cellshape);

                        // draw background
                        var piecenter = cellrect.Center;
                        var pieradius = System.Math.Min(cellrect.Width, cellrect.Height) / 4.0;
                        var piedata = new[] { dp.Value, 1.0 - dp.Value };
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
            foreach (var cell_shape in cell_shapes)
            {
                short shapeid = cell_shape.ID16;
                tb_fmt.Apply(update, shapeid);
                cellfmt.Apply(update, shapeid);
                cellcharfmt.Apply(update, shapeid, 0);
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

            bkfmt.Apply(update, bkshape.ID16);

            update.Execute(page);

            return bkrect.Size;

        }
    }


    public class BarChart : Block
    {
        public IList<DataPoint> DataPoints;
        public VA.Drawing.ColorRGB ValueColor = new VA.Drawing.ColorRGB(0xa0a0a0);
        public VA.Drawing.ColorRGB NonValueColor = new VA.Drawing.ColorRGB(0xffffff);
        double TileHeight = 3.0;
        public BarChart()
        {
            this.DataPoints = new List<DataPoint>();
        }

        public override VA.Drawing.Size Render(RenderContext rc)
        {
            var bkrect = DocUtil.BuildFromUpperLeft(rc.CurrentUpperLeft, new VA.Drawing.Size(rc.PageWidth,TileHeight));

            double bar_width = 0.5;
            double bar_distance = 0.0125;
            double lower_y = bkrect.LowerLeft.Y;
            double maxval = 180.0;
            double label_height = 0.5;

            var margin = new VA.Drawing.Size(0.25, 0.25);
            var inner_ll = bkrect.LowerLeft.Add(margin);
            var inner_ur = bkrect.UpperRight.Subtract(margin);
            var innerrect = new VA.Drawing.Rectangle(inner_ll, inner_ur);

            var bararea_ll = innerrect.LowerLeft.Add(0, label_height);
            var bararea_ur = innerrect.UpperRight;
            var bararea_rect = new VA.Drawing.Rectangle(bararea_ll, bararea_ur);
            
            var xdoc = new VA.DOM.Document();

            var tilerect = xdoc.DrawRectangle(bkrect);
            tilerect.ShapeCells.FillForegnd = rc.TileReal.ToFormula();
            tilerect.ShapeCells.LineWeight = 0;
            tilerect.ShapeCells.LinePattern = 0;
            

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                var dp = this.DataPoints[i];

                double bar_height = dp.Value/maxval*bararea_rect.Height;

                var bar_ll = new VA.Drawing.Point(bar_width + bar_distance, lower_y).Multiply(i,1).Add(margin.Width,margin.Height+label_height);
                var bar_ur = bar_ll.Add(bar_width, bar_height);

                var bar_rect = new VA.Drawing.Rectangle(bar_ll, bar_ur);

                var bar_shape = xdoc.DrawRectangle(bar_rect);
                bar_shape.Text = dp.Value.ToString();
                bar_shape.ShapeCells.LinePattern = 0;
                bar_shape.ShapeCells.LineWeight = 0.0;
                bar_shape.ShapeCells.FillForegnd = ValueColor.ToFormula();
                bar_shape.ShapeCells.VerticalAlign = 0;
                bar_shape.ShapeCells.CharFont = rc.GetFontID(rc.DefaultFont);

                var label_ll = bar_ll.Subtract(0, margin.Height).Add(0,-0.5);
                var label_ur = label_ll.Add(bar_width, label_height);
                var label_rect = new VA.Drawing.Rectangle(label_ll, label_ur);

                var label_shape = xdoc.DrawRectangle(label_rect);
                label_shape.Text = dp.Value.ToString();
                label_shape.ShapeCells.LinePattern = 0;
                label_shape.ShapeCells.LineWeight = 0.0;
                label_shape.ShapeCells.FillPattern= 0;
                label_shape.ShapeCells.VerticalAlign = 0;
                label_shape.ShapeCells.CharFont = rc.GetFontID(rc.DefaultFont);
                label_shape.Text = dp.Label;


            }

            xdoc.Render(rc.Page);

            return bkrect.Size;

        }
    }

}
