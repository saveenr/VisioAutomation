using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    internal class WriteRecordList
    {
        private readonly List<WriteRecord> items;

        int chunksize = -1;

        public WriteRecordList()
        {
            this.items = new List<WriteRecord>();
        }

        public void Clear()
        {
            this.items.Clear();
        }

        public void Add(SidSrc sidsrc, string value)
        {
            CheckForSidSrc();
            var item = new WriteRecord(sidsrc, value);
            this.items.Add(item);
        }

        public void Add(Src coord, string value)
        {
            CheckForSrc();
            var item = new WriteRecord(new SidSrc(-1, coord), value);
            this.items.Add(item);
        }

        private void CheckForSidSrc()
        {
            if (this.chunksize < 0)
            {
                this.chunksize = 4;
            }
            else if (this.chunksize != 4)
            {
                string msg = string.Format("Excpected a src value");
            }
        }

        private void CheckForSrc()
        {
            if (this.chunksize < 0)
            {
                this.chunksize = 3;
            }
            else if (this.chunksize != 3)
            {
                string msg = string.Format("Excpected a sidsrc value");
            }
        }

        public IEnumerable<SidSrc> EnumSidSrcs()
        {
            return this.items.Select(i=>i.SidSrc);
        }

        public IEnumerable<Src> EnumSrcs()
        {
            return this.items.Select(i => i.SidSrc.Src);
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