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
    public class PieSliceChart: GridChart
    {
        public string [] CategoryLabels;
        public DataPoints DataPoints;


        public PieSliceChart(DataPoints dps, string [] cats)
        {
            this.DataPoints = dps;
            this.CategoryLabels = cats;

        }

        public void Draw(Session session)
        {

            var normalized_values = DataPoints.GetNormalizedValues();
            var widths = ConstructPositions(DataPoints.Count(), this.CellWidth , HorizontalSeparation);
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

                cells.FillForegnd = NonValueColor;
                cells.LineColor = LineLightBorder;

            }

            foreach (var shape in slice_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = ValueFillColor;
                cells.LineColor = LineLightBorder;

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = CategoryFillPattern;
                cells.LineWeight = CategoryLineWeight;
                cells.LinePattern = CategoryLinePattern;
            }

            dom.Render(session.Page);
        }
    }


    public class DoughnutSliceChart : GridChart
    {
        public string[] CategoryLabels;
        public DataPoints DataPoints;


        public DoughnutSliceChart(DataPoints dps, string[] cats)
        {
            this.DataPoints = dps;
            this.CategoryLabels = cats;

        }

        public void Draw(Session session)
        {

            var normalized_values = DataPoints.GetNormalizedValues();
            var widths = ConstructPositions(DataPoints.Count(), this.CellWidth, HorizontalSeparation);
            var heights = ConstructPositions(new[] { CategoryLabelHeight, CellHeight }, VerticalSeparation);
            var grid = new GridLayout(widths, heights);

            int catrow = 0;
            int barrow = 2;

            var top_rects = this.SkipOdd(grid.GetRectsInRow(barrow)).ToList();

            var cat_rects = this.SkipOdd(grid.GetRectsInRow(catrow)).ToList();

            var dom = new VA.DOM.Document();
            dom.ResolveAllShapeObjects = true;
            var circle_shapes = new List<VA.DOM.Arc>();
            var slice_shapes = new List<VA.DOM.Arc>();
            for (int i = 0; i < DataPoints.Count; i++)
            {
                var dp = DataPoints[i];
                double start = 0;
                double end = 360 * normalized_values[i];
                double radius = top_rects[i].Width / 2.0;

                var circle_shape = dom.DrawArc(top_rects[i].Center, radius * 0.7, radius, 0, 360);
                circle_shapes.Add(circle_shape);

                var dom_shape = dom.DrawArc(top_rects[i].Center, radius*0.7, radius, start, end);
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

                cells.FillForegnd = NonValueColor;
                cells.LineColor = LineLightBorder;

            }

            foreach (var shape in slice_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillForegnd = ValueFillColor;
                cells.LineColor = LineLightBorder;

            }

            foreach (var shape in cat_shapes)
            {
                var cells = shape.ShapeCells;

                cells.FillPattern = CategoryFillPattern;
                cells.LineWeight = CategoryLineWeight;
                cells.LinePattern = CategoryLinePattern;
            }

            dom.Render(session.Page);
        }
    }

}
