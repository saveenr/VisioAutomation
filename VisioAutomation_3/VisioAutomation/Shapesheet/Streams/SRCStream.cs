using VA=VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public class SRCStream : BaseStream<SRC>
    {
        public SRCStream(int capacity) :
            base(capacity, 3)
        {

        }

        protected override void set_item_at_pos(int index, SRC item)
        {

            this.shortarray.SetItem(index, item.Section, item.Row, item.Cell);
        }

        public static SRCStream FromItems<T>(IList<T> items, System.Func<T, SRC> get_streamitem)
        {
            var s = new SRCStream(items.Count);
            s.Fill(items, get_streamitem);
            return s;
        }

        public static SRCStream FromItems(IList<SRC> items)
        {
            return FromItems(items, c => c);
        }

        public void Add(short section, short row, short cell)
        {
            var streamitem = new VA.ShapeSheet.SRC(section, row, cell);
            this.Add(streamitem);
        }
    }
}