 using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public class ShapeRow<T>: RowBase<T>
    {

        internal ShapeRow(int shape_id, VASS.Internal.ArraySegment<T> cells) :
            base(shape_id, cells)
        {
        }
    }
}