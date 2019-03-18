using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionQueryCache
    {
        List<ShapeCache> list_shapecasheitems;

        public SectionQueryCache()
        {
            this.list_shapecasheitems = new List<ShapeCache>();
        }

        public SectionQueryCache(int capacity)
        {
            this.list_shapecasheitems = new List<ShapeCache>(capacity);
        }

        public void Add(ShapeCache item)
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

        public IEnumerable<ShapeCache> ShapeCacheItems
        {
            get
            {
                return this.list_shapecasheitems;
            }
        }

        public ShapeCache this[int index]
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