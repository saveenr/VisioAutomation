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
            this.SnapXY(x, y, out x, out y);
            return new VA.Drawing.Point(x, y);
        }

        private void SnapXY(double x, double y, out double sx, out double sy)
        {
            sx = this.Round(x, this.SnapSize.Width);
            sy = this.Round(y, this.SnapSize.Height);
        }

        private double Round(double val, double snap_val)
        {
            return Round(val, System.MidpointRounding.AwayFromZero, snap_val);
        }

        /// <summary>
        /// rounds val to the nearest fractional value 
        /// </summary>
        /// <param name="val">the value tp round</param>
        /// <param name="rounding">what kind of rounding</param>
        /// <param name="frac"> round to this value (must be greater than 0.0)</param>
        /// <returns>the rounded value</returns>
        private double Round(double val, System.MidpointRounding rounding, double frac)
        {
            if (frac <= 0)
            {
                throw new System.ArgumentOutOfRangeException("frac", "must be greater than or equal to 0.0");
            }
            double retval = System.Math.Round((val / frac), rounding) * frac;
            return retval;
        }
    }
}