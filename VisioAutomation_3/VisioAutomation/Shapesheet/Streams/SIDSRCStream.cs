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
            this.shortarray.SetItem(index, item.ID, item.Section, item.Row, item.Cell);
        }

        public static SIDSRCStream FromItems<T>(IList<T> items, System.Func<T, SIDSRC> get_streamitem)
        {
            var s = new SIDSRCStream(items.Count);
            s.Fill(items, get_streamitem);
            return s;
        }

        public static SIDSRCStream FromItems(IList<SIDSRC> items)
        {
            return FromItems(items, c => c);
        }

        public void Add(short shapeid, short section, short row, short cell)
        {
            var streamitem = new VA.ShapeSheet.SIDSRC(shapeid, section, row, cell);
            this.Add(streamitem);
        }

        public void Add(short shapeid, SRC src)
        {
            var streamitem = new VA.ShapeSheet.SIDSRC(shapeid, src);
            this.Add(streamitem);
        }
    }
}