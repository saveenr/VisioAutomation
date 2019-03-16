using VASS = VisioAutomation.ShapeSheet;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellOutput<T>: OutputBase
    {
        public VASS.Internal.ArraySegment<T> Cells { get; internal set; }

        internal CellOutput(int shape_id, int count, VASS.Internal.ArraySegment<T> cells) : base(shape_id, count)
        {
            this.Cells = cells;
        }
    }
}