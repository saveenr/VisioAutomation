namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Drop
    {

        public static Microsoft.Office.Interop.Visio.Shape Drop(
            this Microsoft.Office.Interop.Visio.Shape shape,
            Microsoft.Office.Interop.Visio.Master master,
            Core.Point point)
        {
            var output = shape.Drop(master, point.X, point.Y);
            return output;
        }
    }
}