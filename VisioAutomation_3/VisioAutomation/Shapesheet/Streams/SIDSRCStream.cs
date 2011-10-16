using VA=VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class SIDSRCStream : BaseStream<SIDSRC>
    {
        public SIDSRCStream(int capacity) :
            base(capacity, 4)
        {

        }

        protected override void set_item_at_pos(int index, SIDSRC item)
        {
            this.chunked_array.SetItem(index, item.ID, item.Section, item.Row, item.Cell);
        }

        public void Add(short shapeid, short section, short row, short cell)
        {
            var streamitem = new VA.ShapeSheet.SIDSRC(shapeid, section, row, cell);
            this.Add(streamitem);
        }

        public void AddRange(IEnumerable<SIDSRC> items)
        {
            foreach (var sidsrc in items)
            {
                this.Add(sidsrc);
            }
        }

        public void Add(short shapeid, SRC src)
        {
            var streamitem = new VA.ShapeSheet.SIDSRC(shapeid, src);
            this.Add(streamitem);
        }
    }
}