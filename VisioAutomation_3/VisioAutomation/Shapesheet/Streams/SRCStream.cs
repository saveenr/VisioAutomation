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

        protected override void SetItem(int index, SRC item)
        {
            this.chunked_array.SetItem(index, item.Section, item.Row, item.Cell);
        }
    }
}