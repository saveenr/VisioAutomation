using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.GetFormulas(stream);
        }


        public static TResult[] GetResults<TResult>(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.GetResults<TResult>(stream, unitcodes);
        }

        public static int SetFormulas(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.SetFormulas(stream, formulas, flags);
        }

        public static int SetResults(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(master);
            return visobjtarget.SetResults(stream, unitcodes, results, flags);
        }
    }
}