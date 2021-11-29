using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Drop
    {

        public static IVisio.Shape Drop(
            this IVisio.Master master1,
            IVisio.Master master2,
            Core.Point point)
        {
            var output = master1.Drop(master2, point.X, point.Y);
            return output;
        }
    }
}