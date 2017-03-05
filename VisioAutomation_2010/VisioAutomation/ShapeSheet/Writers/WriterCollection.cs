using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    internal class WriterCollection<T>
    {
        private List<WriteRecord> items;

        public WriterCollection()
        {
            this.items = new List<WriteRecord>();
        }

        public void Clear()
        {
            this.items.Clear();
        }

        public void Add(T coord, string value)
        {
            var item = new WriteRecord(coord, value);
            this.items.Add(item);
        }

        public IEnumerable<T> EnumCoords()
        {
            return this.items.Select(i=>i.Coord);
        }

        public object[] BuildValues()
        {
            var array = new object[this.items.Count];
            for (int i = 0; i < this.items.Count; i++)
            {
                array[i] = this.items[i].Value;
            }
            return array;
        }

        public int Count => this.items.Count;

        struct WriteRecord
        {
            public T Coord;
            public string Value;

            public WriteRecord(T coord, string value)
            {
                this.Coord = coord;
                this.Value = value;
            }
        }
    }
}