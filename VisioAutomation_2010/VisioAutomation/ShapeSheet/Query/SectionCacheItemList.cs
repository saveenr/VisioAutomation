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
    }
}