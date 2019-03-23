using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    internal class WriteRecordList<T>
    {
        private readonly List<WriteRecord<T>> items;

        public WriteRecordList()
        {
            this.items = new List<WriteRecord<T>>();
        }

        public void Clear()
        {
            this.items.Clear();
        }

        public void Add(T coord, string value)
        {
            var item = new WriteRecord<T>(coord, value);
            this.items.Add(item);
        }

        public IEnumerable<T> EnumCoords()
        {
            return this.items.Select(i=>i.Coord);
        }

        public object[] BuildValuesArray()
        {
            var array = new object[this.items.Count];
            for (int i = 0; i < this.items.Count; i++)
            {
                array[i] = this.items[i].Value;
            }
            return array;
        }

        public int Count => this.items.Count;
    }
}