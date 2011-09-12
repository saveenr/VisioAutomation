namespace VisioAutomation.Internal
{
    internal static class MathUtil
    {
        public static double Round(double val, double snap_val)
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
        public static double Round(double val, System.MidpointRounding rounding, double frac)
        {
            if (frac <= 0)
            {
                throw new System.ArgumentOutOfRangeException("frac","must be greater than or equal to 0.0");
            }
            double retval = System.Math.Round((val / frac), rounding) * frac;
            return retval;
        }
    }
}