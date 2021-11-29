using System.Collections.Generic;

namespace VisioAutomation.Internal.Extensions
{
    internal static class ExtensionHelpers
    {

        public static IEnumerable<T> ToEnumerable<T>(System.Func<int> get_count, System.Func<int, T> get_item)
        {
            int count = get_count();
            for (int i = 0; i < count; i++)
            {
                var item = get_item(i);
                yield return item;
            }
        }

        public static List<T> ToList<T>(System.Func<int> get_count, System.Func<int, T> get_item)
        {
            int count = get_count();
            var list = new List<T>(count);
            for (int i = 0; i < count; i++)
            {
                var item = get_item(i);
                list.Add(item);
            }
            return list;
        }

        public static string[] _GetFormulas(VisioObjectTarget visobjtarget , ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;

            visobjtarget.DispatchAction(
                (shape) => shape.GetFormulasU(stream.Array, out formulas_sa),
                (master) => master.GetFormulasU(stream.Array, out formulas_sa),
                (page) => page.GetFormulasU(stream.Array, out formulas_sa));

            var formulas = Internal.Helpers.SystemArrayToTypedArray<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] _GetResults<TResult>(VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }
            Internal.Helpers.EnforceValid_ResultType(typeof(TResult));
            
            var flags = Internal.Helpers.GetVisGetSetArgsFromType(typeof(TResult));
            System.Array results_sa = null;

            visobjtarget.DispatchAction(
                (shape) => shape.GetResults(stream.Array, (short)flags, unitcodes, out results_sa),
                (master) => master.GetResults(stream.Array, (short)flags, unitcodes, out results_sa),
                (page) => page.GetResults(stream.Array, (short)flags, unitcodes, out results_sa));


            var results = Internal.Helpers.SystemArrayToTypedArray<TResult>(results_sa);
            return results;
        }

        public static int _SetFormulas(VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            Internal.Helpers.ValidateStreamLengthFormulas(stream, formulas);

            int val = visobjtarget.DispatchFunction(
                (shape) => shape.SetFormulas(stream.Array, formulas, flags),
                (master) => master.SetFormulas(stream.Array, formulas, flags),
                (page) => page.SetFormulas(stream.Array, formulas, flags));

            return val;
        }

        public static int _SetResults(VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.Helpers.ValidateStreamLengthResults(stream, results);

            int val = visobjtarget.DispatchFunction(
                (shape) => shape.SetResults(stream.Array, unitcodes, results, flags),
                (master) => master.SetResults(stream.Array, unitcodes, results, flags),
                (page) => page.SetResults(stream.Array, unitcodes, results, flags));

            return val;
        }
    }
}
