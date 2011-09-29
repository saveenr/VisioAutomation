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

    public struct TableRowGroup
    {
        private readonly int _shapeid;
        private readonly int _count;
        private readonly int _start_row;
        private readonly int _end_row;
        private readonly bool initialized;

        internal TableRowGroup(int sid, int count, int start, int end)
        {
            this._shapeid = sid;
            this._count = count;
            this._start_row = start;
            this._end_row = end;
            this.initialized = true;
        }

        private void CheckInitialized()
        {
            if (!initialized)
            {
                throw new AutomationException("Group not initialized");
            }
        }

        public IEnumerable<int> RowIndices
        {
            get
            {
                this.CheckInitialized();

                if (this.Count < 1)
                {
                    yield break;
                }
                else
                {
                    for (int i = this.StartRow; i <= this.EndRow; i++)
                    {
                        yield return i;
                    }
                }
            }
        }

        public int ShapeID
        {
            get
            {
                this.CheckInitialized();
                return _shapeid;
            }
        }

        public int Count
        {
            get
            {
                this.CheckInitialized();
                return _count;
            }
        }

        public int StartRow
        {
            get
            {
                this.CheckInitialized();
                if (this.Count < 1)
                {
                    throw new AutomationException("Group contains no Rows");
                }
                return _start_row;
            }
        }

        public int EndRow
        {
            get
            {
                this.CheckInitialized();
                if (this.Count < 1)
                {
                    throw new AutomationException("Group contains no Rows");
                }
                return _end_row;
            }
        }
    }
}