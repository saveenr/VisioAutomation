using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionCacheList
    {
        List<List<SectionCache>> items;

        public SectionCacheList()
        {
            this.items = new List<List<SectionCache>>();
        }

        public SectionCacheList(int capacity)
        {
            this.items = new List<List<SectionCache>>(capacity);
        }

        public void AddSectionInfosForShape(List<SectionCache> item)
        {
            this.items.Add(item);
        }

        public int CountShapes => this.items.Count;

        public IEnumerable<List<SectionCache>> EnumSectionInfoForShapes => this.items;

        public List<SectionCache> this[int index] => this.items[index];

    }
}