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
    public class PieSliceChart: Chart
    {
        private double cellwidth = 0.5;
        public double HorizontalSeparation = 0.10;
        public double VerticalSeparation = 0.10;
        public double CellHeight = 0.5;
        public double CategoryLabelHeight = 0.5;
        
        public string [] CategoryLabels;
        public DataPoints DataPoints;

        string def_line_light_border = "rgb(220,220,220)";
        string pie_slice_fill_color = "rgb(240,240,240)";
        string pie_slice_bk_color = "rgb(255,255,255)";
        string cat_fill_path = "0";
        string cat_line_weight = "0.0";
        string cat_line_pattern = "0";

        public PieSliceChart(DataPoints dps, string [] cats)
        {
            this.DataPoints = dps;
            this.CategoryLabels = cats;

        }

        public double CellWidth
        {
            get { return cellwidth; }
            set { cellwidth = value; }
        }

        public void Draw(Session session)
        {

            var normalized_values = DataPoints.GetNormalizedValues();
            var widths = ConstructPositions(DataPoints.Count(), cellwidth, HorizontalSeparation);
            var heights = ConstructPositions(new[] { CategoryLabelHeight, CellHeight }, VerticalSeparation);
            var grid = new GridLayout(widths, heights);

            int catrow = 0;
            int barrow = 2;

            var top_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();

            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveAllShapeObjects = true;
            var circle_shapes = new List<VA.DOM.Oval>();
            var slice_shapes = new List<VA.DOM.PieSlice>();
            for (int i = 0; i < DataPoints.Count; i++)
            {
                var dp = DataPoints[i];
                double start = 0;
                double end = 360*normalized_values[i];
                double radius = top_rects[i].Width/2.0;

                var circle_shape = dom.DrawOval(top_rects[i]);
                circle_shapes.Add(circle_shape);

                var dom_shape = dom.DrawPieSlice(top_rects[i].Center, radius, start, end);
                slice_shapes.Add(dom_shape);
            }

            var cat_shapes = this.DrawRects(dom, cat_rects, session.MasterRectangle);

            for (int i = 0; i < DataPoints.Count; i++)
            {
                slice_shapes[i].Text = DataPoints[i].Text.ToString();
                cat_shapes[i].Text = CategoryLabels[i];
            }

            foreach (var shape in circle_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = pie_slice_bk_color;
                cells.LineColor = def_line_light_border;

            }

            foreach (var shape in slice_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = pie_slice_fill_color;
                cells.LineColor = def_line_light_border;

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = cat_fill_path;
                cells.LineWeight = cat_line_weight;
                cells.LinePattern = cat_line_pattern;
            }
            dom.Render(session.Page);
        }
    }
}
