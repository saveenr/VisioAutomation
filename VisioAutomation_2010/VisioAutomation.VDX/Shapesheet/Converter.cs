namespace VisioAutomation.VDX.ShapeSheet
{
    public static class Converter
    {
        public static double PointsToInches(double points)
        {
            return points/72.0;
        }
    }
}