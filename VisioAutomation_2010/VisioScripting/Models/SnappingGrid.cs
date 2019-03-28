using VisioAutomation.Geometry;

namespace VisioScripting.Models
{
    public class SnappingGrid
    {
        public Size SnapSize { get; }
        
        public SnappingGrid(double w, double h)
        {
            this.SnapSize = new Size(w, h);
        }

        public SnappingGrid( Size size)
        {
            this.SnapSize = size;
        }

        public Size Snap(Size size)
        {
            double x;
            double y;
            this._snap_xy(size.Width,size.Height,out x, out y);
            return new Size(x, y);            
        }

        public Point Snap(Point point)
        {
            double x;
            double y;
            this._snap_xy(point.X,point.Y,out x, out y);
            return new Point(x, y);
        }

        public Point Snap(double x, double y)
        {
            this._snap_xy(x, y, out x, out y);
            return new Point(x, y);
        }

        private void _snap_xy(double x, double y, out double sx, out double sy)
        {
            sx = this._round(x, this.SnapSize.Width);
            sy = this._round(y, this.SnapSize.Height);
        }

        private double _round(double val, double snap_val)
        {
            return this._round(val, System.MidpointRounding.AwayFromZero, snap_val);
        }

        /// <summary>
        /// rounds val to the nearest fractional value 
        /// </summary>
        /// <param name="val">the value tp round</param>
        /// <param name="rounding">what kind of rounding</param>
        /// <param name="frac"> round to this value (must be greater than 0.0)</param>
        /// <returns>the rounded value</returns>
        private double _round(double val, System.MidpointRounding rounding, double frac)
        {
            if (frac <= 0)
            {
                throw new System.ArgumentOutOfRangeException(nameof(frac), "must be greater than or equal to 0.0");
            }
            double retval = System.Math.Round((val / frac), rounding) * frac;
            return retval;
        }
    }
}