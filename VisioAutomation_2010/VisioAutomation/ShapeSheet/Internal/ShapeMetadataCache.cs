using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class ShapeMetadataCache : IEnumerable<ShapeMetadataCacheItem>
    {
        private readonly List<ShapeMetadataCacheItem> _list_shapecacheitems;

        public ShapeMetadataCache(int capacity)
        {
            this._list_shapecacheitems = new List<ShapeMetadataCacheItem>(capacity);
        }

        public void Add(ShapeMetadataCacheItem item)
        {
            this._list_shapecacheitems.Add(item);
        }

        public IEnumerator<ShapeMetadataCacheItem> GetEnumerator()
        {
            return this._list_shapecacheitems.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count
        {
            get { return this._list_shapecacheitems.Count; }
        }

        public int CountCells()
        {
            int n = 0;
            foreach (var shapecacheitem in this._list_shapecacheitems)
            {
                n += shapecacheitem.RowCount * shapecacheitem.ColumnGroup.Count;
            }

            return n;
        }
    }
}