using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class SectionMetadataCache
    {
        private readonly List<ShapeCache> _list;

        public SectionMetadataCache()
        {
            this._list = new List<ShapeCache>();
        }

        public SectionMetadataCache(int capacity)
        {
            this._list = new List<ShapeCache>(capacity);
        }

        public void Add(ShapeCache item)
        {
            this._list.Add(item);
        }

        public int Count
        {
            get { return this._list.Count; }
        }

        public IEnumerable<ShapeCache> ShapeCacheItems
        {
            get { return this._list; }
        }

        public ShapeCache this[int index]
        {
            get { return this._list[index]; }
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