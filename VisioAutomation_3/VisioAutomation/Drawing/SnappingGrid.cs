using System;
using System.Collections.Generic;
using VA=VisioAutomation;

namespace VisioAutomation.Drawing
{
    public class SnappingGrid
    {
        public VA.Drawing.Size SnapSize { get; private set; }
        
        public SnappingGrid(double w, double h)
        {
            this.SnapSize = new VA.Drawing.Size(w, h);
        }

        public SnappingGrid( VA.Drawing.Size size)
        {
            this.SnapSize = size;
        }

        public VA.Drawing.Size Snap(VA.Drawing.Size size)
        {
            double x;
            double y;
            this.SnapXY(size.Width,size.Height,out x, out y);
            return new VA.Drawing.Size(x, y);            
        }

        public VA.Drawing.Point Snap(VA.Drawing.Point point)
        {
            double x;
            double y;
            this.SnapXY(point.X,point.Y,out x, out y);
            return new VA.Drawing.Point(x, y);
        }

        public VA.Drawing.Point Snap(double x, double y)
        {
            double ox;
            double oy;
            this.SnapXY(x, y, out x, out y);
            return new VA.Drawing.Point(x, y);
        }

        private void SnapXY(double x, double y, out double sx, out double sy)
        {
            sx = VA.Internal.MathUtil.Round(x, this.SnapSize.Width);
            sy = VA.Internal.MathUtil.Round(y, this.SnapSize.Height);
        }
    }
}