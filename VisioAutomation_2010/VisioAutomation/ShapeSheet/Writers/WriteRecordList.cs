using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class WriteRecordList
    {
        private readonly List<WriteRecord> items;

        CellCoordinateType coordtype;

        public WriteRecordList(CellCoordinateType type)
        {
            this.items = new List<WriteRecord>();
            this.coordtype = type;
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
            if (this.coordtype != CellCoordinateType.SidSrc)
            {
                string msg = string.Format("Excpected a sidsrc value");
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }

        private void CheckForSrc()
        {
            if (this.coordtype != CellCoordinateType.Src)
            {
                string msg = string.Format("Excpected a src value");
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }

        public Streams.StreamArray BuildSidSrcStream()
        {
            if (this.coordtype != CellCoordinateType.SidSrc)
            {
                string msg = string.Format("writer does not contain sidsrcvalues");
                throw new System.ArgumentOutOfRangeException(msg);
            }
            return Streams.StreamArray.FromSidSrc(this.Count, this.EnumSidSrcs());
        }

        public Streams.StreamArray BuildSrcStream()
        {
            if (this.coordtype != CellCoordinateType.Src)
            {
                string msg = string.Format("writer does not contain srcvalues");
                throw new System.ArgumentOutOfRangeException(msg);
            }
            return Streams.StreamArray.FromSrc(this.Count, this.EnumSrcs());
        }

        private IEnumerable<SidSrc> EnumSidSrcs()
        {
            return this.items.Select(i=>i.SidSrc);
        }

        private IEnumerable<Src> EnumSrcs()
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