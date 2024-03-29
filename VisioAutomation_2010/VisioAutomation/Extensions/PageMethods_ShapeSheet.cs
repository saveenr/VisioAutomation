using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class PageMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this IVisio.Page page,
            ShapeSheet.Streams.StreamArray stream)
        {
            return page.Wrap().GetFormulas(stream);
        }

        public static TResult[] GetResults<TResult>(this IVisio.Page page,
            ShapeSheet.Streams.StreamArray stream,
            object[] unitcodes)
        {
            return page.Wrap().GetResults<TResult>(stream, unitcodes);
        }

        public static int SetFormulas(this IVisio.Page page,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            return page.Wrap().SetFormulas(stream, formulas, flags);
        }

        public static int SetResults(this IVisio.Page page,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            return page.Wrap().SetResults(stream, unitcodes, results, flags);
        }
    }
}