using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Models.Charting
{
    public class PieChart
    {
        public double Radius= 1;
        public double InnerRadius = -1;
        public VA.Drawing.Point Center;
        public List<DataPoint> DataPoints;
 
        public PieChart()
        {
            this.DataPoints = new List<DataPoint>();
        }

        public void Render( IVisio.Page page)
        {
            var values = this.DataPoints.Select(p => p.Value).ToList();
            var shapes = new List<IVisio.Shape>(values.Count);
            if (this.InnerRadius <= 0)
            {
                var slices = VA.Models.Charting.PieSlice.GetSlicesFromValues(this.Center, this.Radius, values);
                foreach (var slice in slices)
                {
                    shapes.Add(slice.Render(page));
                }
            }
            else
            {
                var slices = VA.Models.Charting.PieSlice.GetSlicesFromValues(this.Center, this.InnerRadius, this.Radius, values);
                foreach (var slice in slices)
                {
                    shapes.Add(slice.Render(page));
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
            if (allshapes.Count > 0)
            {
                var app = page.Application;
                var win = app.ActiveWindow;
                win.DeselectAll();
                win.DeselectAll();
                win.Select(shapes, IVisio.VisSelectArgs.visSelect);
                var sel = win.Selection;
                sel.Group();                
            }
        }
    }
}