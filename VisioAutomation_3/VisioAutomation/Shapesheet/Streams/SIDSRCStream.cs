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

        protected override void SetItem(int index, SIDSRC item)
        {
            this.chunked_array.SetItem(index, item.ID, item.Section, item.Row, item.Cell);
        }
    }
}