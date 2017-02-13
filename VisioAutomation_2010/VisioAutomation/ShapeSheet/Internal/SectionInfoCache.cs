using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class SectionInfoCache
    {
        List<List<SectionInfo>> items;

        public SectionInfoCache()
        {
            this.items = new List<List<SectionInfo>>();
        }

        public SectionInfoCache(int capacity)
        {
            this.items = new List<List<SectionInfo>>(capacity);
        }

        public void AddSectionInfosForShape(List<SectionInfo> item)
        {
            this.items.Add(item);
        }

        public int CountShapes => this.items.Count;

        public IEnumerable<List<SectionInfo>> EnumSectionInfoForShapes => this.items;

        public List<SectionInfo>GetSectionInfosForShapeAtIndex(int index) => this.items[index];
    }
}