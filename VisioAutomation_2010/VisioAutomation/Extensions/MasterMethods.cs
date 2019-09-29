using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods
    {
        public static Geometry.Rectangle GetBoundingBox(this IVisio.Master master, IVisio.VisBoundingBoxArgs args)
        {
            var surface = new VisioAutomation.SurfaceTarget(master);
            return surface.GetBoundingBox(args);
        }

        public static IEnumerable<IVisio.Master> ToEnumerable(this IVisio.Masters masters)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToEnumerable(() => masters.Count,
                i => masters[i + 1]);
        }

        public static List<IVisio.Master> ToList(this IVisio.Masters masters)
        {
            return VisioAutomation.Internal.Extensions.ExtensionHelpers.ToList(() => masters.Count,
                i => masters[i + 1]);
        }

        public static string[] GetFormulasU(this IVisio.Master master, ShapeSheet.Streams.StreamArray stream)
        {
            System.Array formulas_sa = null;
            master.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = SurfaceTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this IVisio.Master master, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {

            var flags = SurfaceTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            master.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            var results = SurfaceTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }
    }
}