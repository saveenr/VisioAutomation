namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Drop
    {

        public static Microsoft.Office.Interop.Visio.Shape Drop(
            this Microsoft.Office.Interop.Visio.Master master1,
            Microsoft.Office.Interop.Visio.Master master2,
            Core.Point point)
        {
            var output = master1.Drop(master2, point.X, point.Y);
            return output;
        }
    }
}