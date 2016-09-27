using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Queries.Utilities
{
    public class StreamBuilderSRC: StreamBuilderBase
    {

        public StreamBuilderSRC(int capacity)
            : base(3, capacity)
        {
            
        }

        public void Add(short sec, short row, short cell)
        {
            this.__Add_SRC(sec, row, cell);
        }

        public void Add(SRC src)
        {
            this.__Add_SRC(src.Section, src.Row, src.Cell);
        }

        public static short[] CreateStream(IList<SRC> items)
        {
            var streambuilder = new VisioAutomation.ShapeSheet.Queries.Utilities.StreamBuilderSRC(items.Count);

            foreach (var src in items)
            {
                streambuilder.Add(src);
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