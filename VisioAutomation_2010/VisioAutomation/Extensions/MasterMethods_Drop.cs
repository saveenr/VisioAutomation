using System.Collections.Generic;
using System.Linq;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Extensions
{
    public static class MasterMethods_Drop
    {

        public static IVisio.Shape Drop(
            this IVisio.Master master1,
            IVisio.Master master2,
            Core.Point point)
        {
            var output = master1.Drop(master2, point.X, point.Y);
            return output;
        }

        public static short[] DropManyU(
            this IVisio.Master master,
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

            master.DropManyU(masters_obj_array, xy_array, out outids_sa);

            short[] outids = (short[])outids_sa;
            return outids;
        }
    }
}