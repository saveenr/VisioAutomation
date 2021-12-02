using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream)
        {
            return master.Wrap().GetFormulas(stream);
        }


        public static TResult[] GetResults<TResult>(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            return master.Wrap().GetResults<TResult>(stream, unitcodes);
        }

        public static int SetFormulas(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            return master.Wrap().SetFormulas(stream, formulas, flags);
        }

        public static int SetResults(this IVisio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            return master.Wrap().SetResults(stream, unitcodes, results, flags);
        }
    }
}