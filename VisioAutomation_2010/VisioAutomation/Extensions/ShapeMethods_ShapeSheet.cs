﻿namespace VisioAutomation.Extensions
{
    public static class ShapeMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Shape shape, ShapeSheet.Streams.StreamArray stream)
        {
            System.Array formulas_sa = null;
            shape.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = Core.VisioObjectTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Shape shape, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {
            var flags = Core.VisioObjectTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            shape.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            var results = Core.VisioObjectTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }
    }
}