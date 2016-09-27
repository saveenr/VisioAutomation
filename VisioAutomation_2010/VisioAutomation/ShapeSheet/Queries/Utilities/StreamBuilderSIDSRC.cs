using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Utilities
{
    public class StreamBuilderSIDSRC : StreamBuilderBase
    {

        public StreamBuilderSIDSRC(int capacity) : base(4,capacity)
        {
        }

        public void Add(short shape_id, short sec, short row, short cell)
        {
            this.__Add_SIDSRC(shape_id, sec, row, cell);
        }

        public void Add(short shape_id, SRC src)
        {
            this.__Add_SIDSRC(shape_id, src.Section, src.Row, src.Cell);
        }

        public static short[] CreateStream(IList<SIDSRC> items)
        {
            var streambuilder = new VisioAutomation.ShapeSheet.Queries.Utilities.StreamBuilderSIDSRC(items.Count);

            foreach (var sidsrc in items)
            {
                streambuilder.Add(sidsrc.ShapeID, sidsrc.SRC);
            }

            if (!streambuilder.IsFull)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException();
            }

            var stream = streambuilder.Stream;
            return stream;
        }
    }
}