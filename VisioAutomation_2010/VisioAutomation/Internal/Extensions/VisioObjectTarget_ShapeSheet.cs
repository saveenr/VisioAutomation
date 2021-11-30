using VisioAutomation.Internal;

namespace VisioAutomation.Extensions
{
    internal static class VisioObjectTarget_ShapeSheet
    {
        public static string[] GetFormulas(
            this Internal.VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream)
        {
            if (stream.Array.Length == 0)
            {
                return new string[0];
            }

            System.Array formulas_sa = null;

            if (visobjtarget.Category == Internal.VisioObjectCategory.Shape)
            {
                visobjtarget.Shape.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (visobjtarget.Category == Internal.VisioObjectCategory.Master)
            {
                visobjtarget.Master.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (visobjtarget.Category == Internal.VisioObjectCategory.Page)
            {
                visobjtarget.Page.GetFormulasU(stream.Array, out formulas_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            var formulas = Internal.ShapesheetHelpers.SystemArrayToTypedArray<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(
            this Internal.VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream,
            object[] unitcodes)
        {
            if (stream.Array.Length == 0)
            {
                return new TResult[0];
            }

            Internal.ShapesheetHelpers.EnforceValid_ResultType(typeof(TResult));

            var flags = Internal.ShapesheetHelpers.GetVisGetSetArgsFromType(typeof(TResult));
            System.Array results_sa = null;

            if (visobjtarget.Category == Internal.VisioObjectCategory.Shape)
            {
                visobjtarget.Shape.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            }
            else if (visobjtarget.Category == Internal.VisioObjectCategory.Master)
            {
                visobjtarget.Master.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            }
            else if (visobjtarget.Category == Internal.VisioObjectCategory.Page)
            {
                visobjtarget.Page.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            var results = Internal.ShapesheetHelpers.SystemArrayToTypedArray<TResult>(results_sa);
            return results;
        }

        public static int SetFormulas(
            this VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream,
            object[] formulas, short flags)
        {
            Internal.ShapesheetHelpers.ValidateStreamLengthFormulas(stream, formulas);

            int val = 0;
            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                val = visobjtarget.Shape.SetFormulas(stream.Array, formulas, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                val = visobjtarget.Master.SetFormulas(stream.Array, formulas, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                val = visobjtarget.Page.SetFormulas(stream.Array, formulas, flags);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return val;
        }

        public static int SetResults(
            this VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.ShapesheetHelpers.ValidateStreamLengthResults(stream, results);

            int val = 0;
            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                val = visobjtarget.Shape.SetResults(stream.Array, unitcodes, results, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                val = visobjtarget.Master.SetResults(stream.Array, unitcodes, results, flags);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                val = visobjtarget.Page.SetResults(stream.Array, unitcodes, results, flags);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return val;
        }
    }
}