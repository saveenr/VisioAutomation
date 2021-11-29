using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.ShapeSheet.Writers
{
    public class WriteRecordList
    {
        private readonly List<WriteRecord> _items;

        readonly VisioAutomation.ShapeSheet.Streams.StreamType _streamtype;

        public WriteRecordList(VisioAutomation.ShapeSheet.Streams.StreamType type)
        {
            this._items = new List<WriteRecord>();
            this._streamtype = type;
        }

        public void Clear()
        {
            this._items.Clear();
        }

        public void Add(Core.SidSrc sidsrc, string value)
        {
            _check_for_sidsrc();
            var item = new WriteRecord(sidsrc, value);
            this._items.Add(item);
        }

        public void Add(Core.Src src, string value)
        {
            _check_for_src();
            var item = new WriteRecord(new Core.SidSrc(-1, src), value);
            this._items.Add(item);
        }

        private void _check_for_sidsrc()
        {
            if (this._streamtype != VisioAutomation.ShapeSheet.Streams.StreamType.SidSrc)
            {
                string msg = string.Format("Excpected a sidsrc value");
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }

        private void _check_for_src()
        {
            if (this._streamtype != VisioAutomation.ShapeSheet.Streams.StreamType.Src)
            {
                string msg = string.Format("Excpected a src value");
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }


        public Streams.StreamArray BuildStreamArray(VisioAutomation.ShapeSheet.Streams.StreamType type)
        {
            if (this._streamtype != type)
            {
                string msg = string.Format("writer does not contain {0} values", type.ToString() );
                throw new System.ArgumentOutOfRangeException(msg);
            }

            if (type == VisioAutomation.ShapeSheet.Streams.StreamType.Src)
            {
                var srcs = this._items.Select(i => i.SidSrc.Src);
                return Streams.StreamArray.FromSrc(this.Count, srcs);
            }
            else if (type == VisioAutomation.ShapeSheet.Streams.StreamType.SidSrc)
            {
                var sidsrcs = this._items.Select(i => i.SidSrc);
                return Streams.StreamArray.FromSidSrc(this.Count, sidsrcs);
            }
            else
            {
                throw new System.ArgumentOutOfRangeException();
            }
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