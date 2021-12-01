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
            Microsoft.Office.Interop.Visio.Shape output;

            if (this.Category == VisioObjectCategory.Shape)
            {
                output = this.Shape.Drop(master, point.X, point.Y);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                output = this.Master.Drop(master, point.X, point.Y);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                output = this.Page.Drop(master, point.X, point.Y);
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
            Internal.DropHelpers.ValidateDropManyParams(masters, points);

            if (masters.Count < 1)
            {
                return new short[0];
            }

            // NOTE: DropMany will fail if you pass in zero items to drop
            var masters_obj_array = masters.Cast<object>().ToArray();
            var xy_array = Core.Point.ToDoubles(points).ToArray();

            System.Array outids_sa;

            if (this.Category == VisioObjectCategory.Shape)
            {
                this.Shape.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Category == VisioObjectCategory.Master)
            {
                this.Master.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else if (this.Category == VisioObjectCategory.Page)
            {
                this.Page.DropManyU(masters_obj_array, xy_array, out outids_sa);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }

            short[] outids = (short[]) outids_sa;
            return outids;
        }
    }
}