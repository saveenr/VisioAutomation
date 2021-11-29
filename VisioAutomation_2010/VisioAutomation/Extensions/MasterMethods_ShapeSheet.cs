namespace VisioAutomation.Extensions
{
    public static class MasterMethods_ShapeSheet
    {
        public static string[] GetFormulasU(this Microsoft.Office.Interop.Visio.Master master, ShapeSheet.Streams.StreamArray stream)
        {
            System.Array formulas_sa = null;
            master.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = Core.VisioObjectTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this Microsoft.Office.Interop.Visio.Master master, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {

            var flags = Core.VisioObjectTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            master.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            var results = Core.VisioObjectTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }
    }
}