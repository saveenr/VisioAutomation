using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class MultiSectionQueryCache
    {
        List<ShapeCacheItemList> list_shapecasheitems;

        public MultiSectionQueryCache()
        {
            this.list_shapecasheitems = new List<ShapeCacheItemList>();
        }

        public MultiSectionQueryCache(int capacity)
        {
            this.list_shapecasheitems = new List<ShapeCacheItemList>(capacity);
        }

        public void AddSectionInfosForShape(ShapeCacheItemList item)
        {
            this.list_shapecasheitems.Add(item);
        }

        public int Count
        {
            get
            {
                return this.list_shapecasheitems.Count;
            }
        }

        public IEnumerable<ShapeCacheItemList> ShapeCacheItems
        {
            get
            {
                return this.list_shapecasheitems;
            }
        }

        public ShapeCacheItemList this[int index]
        {
            get
            {
                return this.list_shapecasheitems[index];
            }
        }

        public int CountCells()
        {
            // Count the cells not in sections
            int count = 0;
            foreach (var section_info in this.ShapeCacheItems)
            {
                count += section_info.CountCells();
            }

            return count;
        }
    }
}