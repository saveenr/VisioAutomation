using System.Collections.Generic;
using System.Collections;

namespace VisioAutomation.ShapeSheet.Data
{
    public class TableRowGroupList : IEnumerable<TableRowGroup>
    {
        public List<TableRowGroup> items;

        public TableRowGroupList()
        {
            this.items = new List<TableRowGroup>();
        }

        public void Add(TableRowGroup g)
        {
            this.items.Add(g);
        }

        public TableRowGroup this[int index]
        {
            get { return this.items[index]; }
        }

        public int Count
        {
            get { return this.items.Count; }
        }

        public IEnumerator<TableRowGroup> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
        }
    }
}