using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class SectionMetadataCache
    {
        private readonly List<ShapeMetadataCache> _list;

        public SectionMetadataCache()
        {
            this._list = new List<ShapeMetadataCache>();
        }

        public SectionMetadataCache(int capacity)
        {
            this._list = new List<ShapeMetadataCache>(capacity);
        }

        public void Add(ShapeMetadataCache item)
        {
            this._list.Add(item);
        }

        public int Count
        {
            get { return this._list.Count; }
        }

        public IEnumerable<ShapeMetadataCache> ShapeCacheItems
        {
            get { return this._list; }
        }

        public ShapeMetadataCache this[int index]
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