namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Draw
    {

        public static Microsoft.Office.Interop.Visio.Shape DrawLine(this Microsoft.Office.Interop.Visio.Master master, Core.Point p1, Core.Point p2)
        {
            var shape = master.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return shape;
        }

        public static Microsoft.Office.Interop.Visio.Shape DrawQuarterArc(
            this Microsoft.Office.Interop.Visio.Master master,
            Core.Point p0,
            Core.Point p1,
            Microsoft.Office.Interop.Visio.VisArcSweepFlags flags)
        {
            var s = master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            return s;
        }
    }
}