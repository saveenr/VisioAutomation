using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    public class CellQueryResults<T>: RowList<T>
    {
        internal CellQueryResults(int capacity) : base(capacity)
        {

        }
    }
}