using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public class RowBase<T>
    {
        public int ShapeID { get; private set; }
        public readonly VASS.Internal.ArraySegment<T> Cells;

        internal RowBase(int shapeid, VASS.Internal.ArraySegment<T>  cells)
        {
            this.ShapeID = shapeid;
            this.Cells = cells;
        }
    }
}