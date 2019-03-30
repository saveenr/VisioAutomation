using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class WriteRecordList
    {
        private readonly List<WriteRecord> _items;

        readonly CellCoordinateType _coordtype;

        public WriteRecordList(CellCoordinateType type)
        {
            this._items = new List<WriteRecord>();
            this._coordtype = type;
        }

        public void Clear()
        {
            this._items.Clear();
        }

        public void Add(SidSrc sidsrc, string value)
        {
            _check_for_sidsrc();
            var item = new WriteRecord(sidsrc, value);
            this._items.Add(item);
        }

        public void Add(Src coord, string value)
        {
            _check_for_src();
            var item = new WriteRecord(new SidSrc(-1, coord), value);
            this._items.Add(item);
        }

        private void CheckForSidSrc()
        {
            if (this._coordtype != CellCoordinateType.SidSrc)
            {
                string msg = string.Format("Excpected a sidsrc value");
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }

        private void _check_for_src()
        {
            if (this._coordtype != CellCoordinateType.Src)
            {
                string msg = string.Format("Excpected a src value");
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }

        public Streams.StreamArray BuildSidSrcStream()
        {
            if (this._coordtype != CellCoordinateType.SidSrc)
            {
                string msg = string.Format("writer does not contain sidsrcvalues");
                throw new System.ArgumentOutOfRangeException(msg);
            }

            var sidsrcs = this._items.Select(i => i.SidSrc);
            return Streams.StreamArray.FromSidSrc(this.Count, sidsrcs);
        }

        public Streams.StreamArray BuildSrcStream()
        {
            if (this._coordtype != CellCoordinateType.Src)
            {
                string msg = string.Format("writer does not contain srcvalues");
                throw new System.ArgumentOutOfRangeException(msg);
            }

            var srcs = this._items.Select(i => i.SidSrc.Src);
            return Streams.StreamArray.FromSrc(this.Count, srcs);
        }

        public object[] BuildValuesArray()
        {
            var array = new object[this._items.Count];
            for (int i = 0; i < this._items.Count; i++)
            {
                array[i] = this._items[i].Value;
            }
            return array;
        }

        public int Count => this._items.Count;
    }
}