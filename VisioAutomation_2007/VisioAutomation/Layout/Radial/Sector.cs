using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Radial
{
    public struct Sector
    {
        public double StartAngle { get; private set; }
        public double EndAngle { get; private set; }

        public Sector(double start, double end) :
            this()
        {
            this.StartAngle = start;
            this.EndAngle = end;
        }

        public static List<Sector> GetSectorsFromValues(IList<double> values)
        {
            double sectors = values.Sum();
            var slices = new List<Sector>(values.Count);
            double start_angle = 0;
            foreach (int i in Enumerable.Range(0, values.Count))
            {
                double cur_val = values[i];
                double cur_val_norm = cur_val / sectors;
                double cur_angle = cur_val_norm * System.Math.PI * 2.0;
                double end_angle = start_angle + cur_angle;

                var ps = new VA.Layout.Radial.Sector(start_angle, end_angle);
                slices.Add(ps);

                start_angle += cur_angle;
            }
            return slices;
        }

        public double Angle
        {
            get { return this.EndAngle - this.StartAngle; }
        }
    }
}