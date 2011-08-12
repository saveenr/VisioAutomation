using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VA=VisioAutomation;

namespace VisioAutomation.Layout
{
    public class PieSlice
    {
        public VA.Drawing.Point Center { get; set; }
        public double Radius { get; set; }
        public double StartAngle { get; set; }
        public double EndAngle { get; set; }

        public PieSlice(VA.Drawing.Point center, double radius, double start, double end)
        {
            this.Center = center;
            this.Radius = radius;
            this.StartAngle = start;
            this.EndAngle = end;
        }
    }
}
