using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionCache
    {
        List<ShapeCacheItemList> items;

        public SectionCache()
        {
            this.items = new List<ShapeCacheItemList>();
        }

        public SectionCache(int capacity)
        {
            this.items = new List<ShapeCacheItemList>(capacity);
        }

        public void AddSectionInfosForShape(ShapeCacheItemList item)
        {
            this.items.Add(item);
        }

        public int CountShapes
        {
            get
            {
                return this.items.Count;
            }
        }

        public IEnumerable<ShapeCacheItemList> ShapeCacheItems
        {
            get
            {
                return this.items;
            }
        }

        public ShapeCacheItemList this[int index]
        {
            get
            {
                return this.items[index];
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