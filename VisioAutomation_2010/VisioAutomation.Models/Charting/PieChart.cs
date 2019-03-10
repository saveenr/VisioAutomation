using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Charting
{
    public class PieChart
    {
        public double Radius= 1;
        public double InnerRadius = -1;
        public VisioAutomation.Geometry.Point Center;
        public DataPointList DataPoints;
 
        public PieChart(VisioAutomation.Geometry.Point center, double radius)
        {
            this.DataPoints = new DataPointList();
            this.Center = center;
            this.Radius = radius;
        }

        public PieChart(VisioAutomation.Geometry.Rectangle rect)
        {
            var center = rect.Center;
            var radius = System.Math.Min(rect.Width,rect.Height)/2.0;
            this.DataPoints = new DataPointList();
            this.Center = center;
            this.Radius = radius;
        }

        public void Render( IVisio.Page page)
        {
            var values = this.DataPoints.Select(p => p.Value).ToList();
            var shapes = new List<IVisio.Shape>(values.Count);
            if (this.InnerRadius <= 0)
            {
                var slices = PieSlice.GetSlicesFromValues(this.Center, this.Radius, values);
                foreach (var slice in slices)
                {
                    var rendered_shape = slice.Render(page);
                    shapes.Add(rendered_shape);
                }
            }
            else
            {
                var slices = PieSlice.GetSlicesFromValues(this.Center, this.InnerRadius, this.Radius, values);
                foreach (var slice in slices)
                {
                    var rendered_shape = slice.Render(page);
                    shapes.Add(rendered_shape);
                }
            }

            for (int i = 0; i < this.DataPoints.Count; i++)
            {
                var dp = this.DataPoints[i];
                var shape = shapes[i];

                dp.VisioShape = shape;
                if (dp.Label != null)
                {
                    if (dp.LabelFormat != null)
                    {
                        string formatted_label = string.Format(dp.Label, dp.Label);
                        dp.VisioShape.Text = formatted_label;
                    }
                    else
                    {
                        dp.VisioShape.Text = dp.Label;
                    }
                }
            }

            var allshapes = this.DataPoints.Select(dp => dp.VisioShape).Where(s => s != null).ToList();
            ChartUtil.GroupShapesIfNeeded(page, allshapes);
        }
    }
}