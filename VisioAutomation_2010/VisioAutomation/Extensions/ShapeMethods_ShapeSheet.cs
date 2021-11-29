namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Shape shape, ShapeSheet.Streams.StreamArray stream)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.GetFormulas(stream);
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Shape shape, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.GetResults<TResult>(stream, unitcodes);
        }

        public static int SetFormulas(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.SetFormulas(stream, formulas, flags);
        }

        public static int SetResults(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(shape);
            return visobjtarget.SetResults(stream, unitcodes, results, flags);
        }
    }
}