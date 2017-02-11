using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    internal class WriteRecords<T>
    {
        private readonly List<WriteRecord<T>> Records;

        public WriteRecords()
        {
            this.Records = new List<WriteRecord<T>>();
        } 

        public void Clear()
        {
            this.Records.Clear();
        }

        public void Add(SRC src, T value)
        {
            var rec = new WriteRecord<T>(src, value);
            this.Records.Add(rec);
        }

        public void Add(SIDSRC sidsrc, T value)
        {
            var rec = new WriteRecord<T>(sidsrc, value);
            this.Records.Add(rec);
        }

        public int Count => this.Records.Count;

        public IEnumerable<WriteRecord<T>> Enum(CoordType type)
        {
            return this.Records.Where(i => i.Type == type);
        }
    }
}