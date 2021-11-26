using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods
    {
        public static Core.Rectangle GetBoundingBox(this IVisio.Master master, IVisio.VisBoundingBoxArgs args)
        {
            double bbx0, bby0, bbx1, bby1;
            master.BoundingBox((short)args, out bbx0, out bby0, out bbx1, out bby1);
            var r = new Core.Rectangle(bbx0, bby0, bbx1, bby1);
            return r;
        }

        public static IEnumerable<IVisio.Master> ToEnumerable(this IVisio.Masters masters)
        {
            return Internal.Extensions.ExtensionHelpers.ToEnumerable(() => masters.Count,
                i => masters[i + 1]);
        }

        public static List<IVisio.Master> ToList(this IVisio.Masters masters)
        {
            return Internal.Extensions.ExtensionHelpers.ToList(() => masters.Count,
                i => masters[i + 1]);
        }

        public static string[] GetFormulasU(this IVisio.Master master, ShapeSheet.Streams.StreamArray stream)
        {
            System.Array formulas_sa = null;
            master.GetFormulasU(stream.Array, out formulas_sa);
            var formulas = Core.SurfaceTarget.system_array_to_typed_array<string>(formulas_sa);
            return formulas;
        }

        public static TResult[] GetResults<TResult>(this IVisio.Master master, ShapeSheet.Streams.StreamArray stream, object[] unitcodes)
        {

            var flags = Core.SurfaceTarget._type_to_vis_get_set_args(typeof(TResult));
            System.Array results_sa = null;
            master.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            var results = Core.SurfaceTarget.system_array_to_typed_array<TResult>(results_sa);
            return results;
        }

        public static IVisio.Shape DrawLine(this IVisio.Master master, Core.Point p1, Core.Point p2)
        {
            var shape = master.DrawLine(p1.X, p1.Y, p2.X, p2.Y);
            return shape;
        }

        public static IVisio.Shape DrawQuarterArc(
            this IVisio.Master master,
            Core.Point p0,
            Core.Point p1,
            IVisio.VisArcSweepFlags flags)
        {
            var s = master.DrawQuarterArc(p0.X, p0.Y, p1.X, p1.Y, flags);
            return s;
        }
    }
}