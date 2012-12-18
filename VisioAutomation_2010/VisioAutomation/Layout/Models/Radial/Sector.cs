using System;
using System.Collections.Generic;
using System.Linq;
using VisioAutomation.Drawing;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;

namespace VisioAutomation.Layout.Models.Radial
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

        public double Angle
        {
            get { return this.EndAngle - this.StartAngle; }
        }
    }
}