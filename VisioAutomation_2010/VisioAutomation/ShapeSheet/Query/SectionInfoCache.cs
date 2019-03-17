using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class SectionInfoCache
    {
        List<List<SectionCacheInfo>> items;

        public SectionInfoCache()
        {
            this.items = new List<List<SectionCacheInfo>>();
        }

        public SectionInfoCache(int capacity)
        {
            this.items = new List<List<SectionCacheInfo>>(capacity);
        }

        public void AddSectionInfosForShape(List<SectionCacheInfo> item)
        {
            this.items.Add(item);
        }

        public int CountShapes => this.items.Count;

        public IEnumerable<List<SectionCacheInfo>> EnumSectionInfoForShapes => this.items;

        public List<SectionCacheInfo> this[int index] => this.items[index];

    }
}