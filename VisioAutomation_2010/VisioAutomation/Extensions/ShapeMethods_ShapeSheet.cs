namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Shape shape, ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;
            shape.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = Core.VisioObjectTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Shape shape, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }
            Internal.TempHelper._enforce_valid_result_type(typeof(TResult));


            var flags = Core.VisioObjectTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            shape.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            var results = Core.VisioObjectTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }
        public static int SetFormulas(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            int val = shape.SetFormulas(stream.Array, formulas, flags);
            return val;
        }

        public static int SetResults(this Microsoft.Office.Interop.Visio.Shape shape,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.TempHelper.ValidateStreamLengthResults(stream, results);

            int val = shape.SetResults(stream.Array, unitcodes, results, flags);
            return val;
        }
    }
}