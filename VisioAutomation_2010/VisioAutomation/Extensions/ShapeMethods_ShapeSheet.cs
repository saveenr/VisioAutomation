namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream)
        {
            return shape.Wrap().GetFormulas(stream);
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            return shape.Wrap().GetResults<TResult>(stream, unitcodes);
        }

        public static int SetFormulas(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            return shape.Wrap().SetFormulas(stream, formulas, flags);
        }

        public static int SetResults(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            return shape.Wrap().SetResults(stream, unitcodes, results, flags);
        }
    }
}