using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_Drop
    {

        public static IVisio.Shape Drop(
            this IVisio.Shape shape,
            IVisio.Master master,
            Core.Point point)
        {
            var output = shape.Drop(master, point.X, point.Y);
            return output;
        }
    }
}