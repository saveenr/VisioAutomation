using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{

    internal class ShapeCacheItemList : List<ShapeCacheItem>
    {
        public ShapeCacheItemList()
        {

        }

        public ShapeCacheItemList(int capacity) : base(capacity)
        {

        }

        public int CountCells()
        {
            int n = 0;
            foreach (var shapecacheitem in this)
            {
                n += shapecacheitem.RowCount * shapecacheitem.SectionQuery.Columns.Count;
            }
            return n;
        }
    }
}