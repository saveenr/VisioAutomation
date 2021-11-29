namespace VisioAutomation.Extensions
{
    public static class MasterMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Master master,
            ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;
            master.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = Internal.Helpers.SystemArrayToTypedArray<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }
            Internal.Helpers.EnforceValid_ResultType(typeof(TResult));


            var flags = Internal.Helpers.GetVisGetSetArgsFromType(typeof(TResult));
            System.Array results_sa = null;
            master.GetResults(stream.Array, (short) flags, unitcodes, out results_sa);
            var results = Internal.Helpers.SystemArrayToTypedArray<TResult>(results_sa);
            return results;
        }

        public static int SetFormulas(this Microsoft.Office.Interop.Visio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            Internal.Helpers.ValidateStreamLengthFormulas(stream, formulas);

            int val = master.SetFormulas(stream.Array, formulas, flags);
            return val;
        }

        public static int SetResults(this Microsoft.Office.Interop.Visio.Master master,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.Helpers.ValidateStreamLengthResults(stream, results);

            int val = master.SetResults(stream.Array, unitcodes, results, flags);
            return val;
        }
    }
}