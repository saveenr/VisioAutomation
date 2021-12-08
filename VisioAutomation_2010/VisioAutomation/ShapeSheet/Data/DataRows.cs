using System.Collections.Generic;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Core
{
    public class BasicList<T> : IEnumerable<T>
    {
        private readonly List<T> _list;

        internal BasicList()
        {
            this._list = new List<T>();
        }

        internal BasicList(int capacity)
        {
            this._list = new List<T>(capacity);
        }
        public IEnumerator<T> GetEnumerator()
        {
            return this._list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(T r)
        {
            this._list.Add(r);
        }

        public void AddRange(IEnumerable<T> rows)
        {
            this._list.AddRange(rows);
        }

        public int Count
        {
            get
            {
                return this._list.Count;
            }
        }

        public T this[int index]
        {
            get { return this._list[0]; }
            //set { /* set the specified index to value here */ }
        }
    }
}

namespace VisioAutomation.ShapeSheet.Data
{


    public class DataRows<T> : VisioAutomation.Core.BasicList<DataRow<T>>
    {
        // Simple list of Rows


        public readonly int ShapeID;
        public readonly IVisio.VisSectionIndices SectionIndex;

        internal DataRows(int capacity) : base(capacity)
        {
            this.ShapeID = -1;
            this.SectionIndex = IVisio.VisSectionIndices.visSectionInval;
        }

        internal DataRows(int capacity, int shapeid, IVisio.VisSectionIndices section_index) : base (capacity)
        {
            this.ShapeID = shapeid;
            this.SectionIndex = section_index;
        }
    }
}