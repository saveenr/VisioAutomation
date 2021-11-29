namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Draw
    {
        public static Microsoft.Office.Interop.Visio.Shape DrawLine(
            this Microsoft.Office.Interop.Visio.Shape shape,
            Core.Point p1, Core.Point p2)
        {
            var s = shape.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return s;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawQuarterArc(
            this Microsoft.Office.Interop.Visio.Shape shape,
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
        {
            var s = shape.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            return s;
        }
    }
}