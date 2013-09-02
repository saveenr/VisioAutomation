using System;
using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Models.Radial
{
    public class PieChart
    {
        public double Radius= 1;
        public double InnerRadius = -1;
        public VA.Drawing.Point Center;
        public List<double> Values;
        public IList<string> Labels; 

        public PieChart(IList<double> values)
        {
            this.Values = values.ToList();
        }

        public List<IVisio.Shape> Render( IVisio.Page page)
        {
            if (this.InnerRadius <= 0)
            {
                var slices = VA.Layout.Models.Radial.PieSlice.GetSlicesFromValues(this.Center, this.Radius, this.Values);
                var shapes = new List<IVisio.Shape>(slices.Count);
                foreach (var slice in slices)
                {
                    var shape = slice.Render(page);
                    shapes.Add(shape);
                    this.SetText(shapes);
                }
                return shapes;
            }
            else
            {
                var slices = VA.Layout.Models.Radial.DoughnutSlice.GetSlicesFromValues(this.Center, this.InnerRadius, this.Radius, this.Values);
                var shapes = new List<IVisio.Shape>(slices.Count);
                foreach (var slice in slices)
                {
                    var shape = slice.Render(page);
                    shapes.Add(shape);
                    this.SetText(shapes);
                }
                return shapes;
            }
        }

        private void SetText(List<IVisio.Shape> shapes)
        {
            if (this.Labels == null)
            {
                return;
            }

            for (int i = 0; i < shapes.Count; i++)
            {
                if (i < this.Labels.Count)
                {
                    string label = this.Labels[i];
                    var shape = shapes[i];
                    shape.Text = label;
                }
            }
        }
    }
}