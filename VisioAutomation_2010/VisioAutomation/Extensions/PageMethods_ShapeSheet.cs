namespace VisioAutomation.Extensions
{
    public static class PageMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return VisioAutomation.Extensions.ExtensionHelpers.GetFormulas(visobjtarget, stream);
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream,
            object[] unitcodes)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.GetResults<TResult>(stream, unitcodes);
        }

        public static int SetFormulas(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return visobjtarget.SetFormulas(stream, formulas, flags);

        }

        public static int SetResults(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            var visobjtarget = new VisioAutomation.Internal.VisioObjectTarget(page);
            return VisioAutomation.Extensions.ExtensionHelpers.SetResults(visobjtarget, stream, unitcodes, results, flags);

        }
    }
}
