using System.Collections.Generic;
using System.Linq;
using IVisio=Microsoft.Office.Interop.Visio;

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

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                visobjtarget.Shape.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                visobjtarget.Master.GetFormulasU(stream.Array, out formulas_sa);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                visobjtarget.Page.GetFormulasU(stream.Array, out formulas_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
            
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

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                visobjtarget.Shape.GetResults(stream.Array, (short) flags, unitcodes, out results_sa);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                visobjtarget.Master.GetResults(stream.Array, (short) flags, unitcodes, out results_sa);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                visobjtarget.Page.GetResults(stream.Array, (short)flags, unitcodes, out results_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            var results = Internal.Helpers.SystemArrayToTypedArray<TResult>(results_sa);
            return results;
        }

        public static int _SetFormulas(VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream, object[] formulas, short flags)
        {
            Internal.Helpers.ValidateStreamLengthFormulas(stream, formulas);

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

        public static int _SetResults(VisioObjectTarget visobjtarget,
            ShapeSheet.Streams.StreamArray stream, object[] unitcodes, object[] results, short flags)
        {
            Internal.Helpers.ValidateStreamLengthResults(stream, results);

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


        public static IVisio.Shape _Drop(
            VisioObjectTarget visobjtarget,
            IVisio.Master master,
            Core.Point point)
        {
            IVisio.Shape output;

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                output = visobjtarget.Shape.Drop(master, point.X, point.Y);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                output = visobjtarget.Master.Drop(master, point.X, point.Y);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                output = visobjtarget.Page.Drop(master, point.X, point.Y);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            return output;
        }

        public static short[] _DropManyU(
            VisioObjectTarget visobjtarget,
            IList<IVisio.Master> masters,
            IEnumerable<Core.Point> points)
        {
            Internal.Helpers.ValidateDropManyParams(masters, points);


            if (masters.Count < 1)
            {
                return new short[0];
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = Core.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            if (visobjtarget.Category == VisioObjectCategory.Shape)
            {
                visobjtarget.Shape.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Master)
            {
                visobjtarget.Master.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (visobjtarget.Category == VisioObjectCategory.Page)
            {
                visobjtarget.Page.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            short[] outids = (short[])outids_sa;
            return outids;
        }
    }
}
