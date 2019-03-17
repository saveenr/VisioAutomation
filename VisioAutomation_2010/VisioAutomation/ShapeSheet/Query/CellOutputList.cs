using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellOutputList<T> : List<CellOutput<T>>
    {
        internal CellOutputList() : base()
        {
        }

        internal CellOutputList(int capacity) : base(capacity)
        {
        }
    }
}