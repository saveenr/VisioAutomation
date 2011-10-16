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

        public void AddRange(IEnumerable<SRC> items)
        {
            foreach (var src in items)
            {
                this.Add(src);
            }
        }

        public void Add(short section, short row, short cell)
        {
            var streamitem = new VA.ShapeSheet.SRC(section, row, cell);
            this.Add(streamitem);
        }
    }
}