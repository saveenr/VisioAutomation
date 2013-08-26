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
                }
                return shapes;
            }
        }
    }
}