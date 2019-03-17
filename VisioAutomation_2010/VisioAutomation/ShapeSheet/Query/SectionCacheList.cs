using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Query
{
    internal class MSCache
    {
        List<LISTSECTIONCLASS> items;

        public MSCache()
        {
            this.items = new List<LISTSECTIONCLASS>();
        }

        public MSCache(int capacity)
        {
            this.items = new List<LISTSECTIONCLASS>(capacity);
        }

        public void AddSectionInfosForShape(LISTSECTIONCLASS item)
        {
            this.items.Add(item);
        }

        public int CountShapes => this.items.Count;

        public IEnumerable<LISTSECTIONCLASS> EnumSectionInfoForShapes => this.items;

        public LISTSECTIONCLASS this[int index] => this.items[index];

    }

    internal class LISTSECTIONCLASS : List<SectionCache>
    {
        public LISTSECTIONCLASS()
        {

        }

        public LISTSECTIONCLASS(int capacity) : base(capacity)
        {

        }
    }
}