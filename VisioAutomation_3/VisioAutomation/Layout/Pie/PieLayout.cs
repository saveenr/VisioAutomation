using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Pie
{
    public class PieLayout
    {
        private List<PieSlice> _slices;
        public VA.Drawing.Point Center;
        private double _radius;
        public bool DrawBackground { get; set; }
        private IVisio.Shape _backgroundVisioShape;

        public PieLayout()
        {
            this.Center = new VA.Drawing.Point(0,0);
            this._radius = 1.0;
            this._slices = new List<PieSlice>();
        }

        public List<PieSlice> Slices
        {
            get { return _slices; }
        }

        public double Radius
        {
            get { return _radius; }
            set
            {
                if (value <= 0)
                {
                    throw new System.ArgumentOutOfRangeException("value");
                }
                _radius = value;
            }
        }

        public Shape BackgroundCircle
        {
            get { return _backgroundVisioShape; }
        }

        public void Add(PieSlice slice)
        {
            this._slices.Add(slice);
        }

        public PieSlice Add(double value, string text)
        {
            var slice = new PieSlice();
            slice.Value = value;
            slice.Text = text;
            this._slices.Add(slice);
            return slice;
        }

        public void Render(IVisio.Page page)
        {
            double sum = this._slices.Select(s => s.Value).Sum();
            var shapes = new List<IVisio.Shape>();
            double start_angle = 0.0;

            // Draw each slice
            for (int i = 0; i < this.Slices.Count; i++)
            {
                var slice = this.Slices[i];
                double cur_val = slice.Value;
                double cur_val_norm = cur_val / sum;
                double cur_angle = cur_val_norm * System.Math.PI * 2.0;
                double end_angle = start_angle + cur_angle;


                slice.StartAngle = start_angle;
                slice.EndAngle = end_angle;

                var ps = new VA.Layout.Radial.PieSlice(this.Center, start_angle, end_angle, this.Radius);
                var shape = ps.Render(page);
                start_angle += cur_angle;

                shapes.Add(shape);
            }


            for (int i = 0; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var slice = this._slices[i];
                slice.VisioShape = shape;
                if (!string.IsNullOrEmpty(slice.Text))
                {
                    shape.Text = slice.Text;
                }
            }

            if (this.DrawBackground)
            {
                var ll = this.Center.Subtract(this.Radius, this.Radius);
                var ur = this.Center.Add(this.Radius, this.Radius);
                var rect = new VA.Drawing.Rectangle(ll, ur);
                this._backgroundVisioShape = page.DrawOval(rect);
            }
        }
    }
}