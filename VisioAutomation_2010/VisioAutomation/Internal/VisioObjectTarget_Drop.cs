using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.Internal
{
    internal readonly partial struct VisioObjectTarget
    {
        public Microsoft.Office.Interop.Visio.Shape Drop(
            Microsoft.Office.Interop.Visio.Master master,
            Core.Point point)
        {


            var visobjtarget = this;


            Microsoft.Office.Interop.Visio.Shape output;

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

        public short[] DropManyU(
            IList<Microsoft.Office.Interop.Visio.Master> masters,
            IEnumerable<Core.Point> points)
        {

            var visobjtarget = this;

            Internal.DropHelpers.ValidateDropManyParams(masters, points);
            
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