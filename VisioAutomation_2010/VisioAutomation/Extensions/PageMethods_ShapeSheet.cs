namespace VisioAutomation.Extensions
{
    public static class PageMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;
            page.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = Core.VisioObjectTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream,
            object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }
            Internal.TempHelper._enforce_valid_result_type(typeof(TResult));


            var flags = Core.VisioObjectTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            page.GetResults(stream.Array, (short) flags, unitcodes, out results_sa);
            var results = Core.VisioObjectTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }

        public static int SetFormulas(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            Internal.TempHelper.ValidateStreamLengthFormulas(stream, formulas);

            int val = page.SetFormulas(stream.Array, formulas, flags);
            return val;
        }

        public static int SetResults(this Microsoft.Office.Interop.Visio.Page page,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.TempHelper.ValidateStreamLengthResults(stream, results);

            int val = page.SetResults(stream.Array, unitcodes, results, flags);
            return val;
        }
    }
}
