using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public struct StreamArray
    {
        public readonly short[] Array;
        private readonly int ChunkSize;
        public readonly int Count;
        public readonly StreamType Type;

        internal StreamArray(short[] array, StreamType type, int count)
        {
            this.Array = array ?? throw new System.ArgumentNullException(nameof(array));
            this.Type = type;
            this.ChunkSize = type == StreamType.SidSrc ? 4 : 3;
            this.Count = count;

            if (array.Length % this.ChunkSize != 0)
            {
                string msg = string.Format("Array length must be a multiple of {0}", this.ChunkSize);
                throw new Exceptions.InternalAssertionException(msg);
            }

            if (count * this.ChunkSize != array.Length)
            {
                string msg =
                    string.Format(
                        "Array length does not match the number of cells {0} multiplied by the chunk size {1}",
                        this.Count, this.ChunkSize);
                throw new Exceptions.InternalAssertionException(
                    "Count does not match the number of short elements in the array");
            }

            this.Count = count;
        }


        public static StreamArray FromSidSrc(IList<Core.SidSrc> sidsrcs)
        {
            return FromSidSrc(sidsrcs.Count, sidsrcs);
        }

        public static StreamArray FromSidSrc(int numcells, IEnumerable<Core.SidSrc> sidsrcs)
        {
            var num_shorts = numcells * 4;
            var array = new short[num_shorts];
            var stream = new StreamArray(array, StreamType.SidSrc, numcells);

            int i = 0;
            int j = 0;
            foreach (var sidsrc in sidsrcs)
            {
                if (j >= numcells)
                {
                    break;
                }

                array[i++] = sidsrc.ShapeID;
                array[i++] = sidsrc.Src.Section;
                array[i++] = sidsrc.Src.Row;
                array[i++] = sidsrc.Src.Cell;
                j++;
            }

            return stream;
        }

        public static StreamArray FromSrc(IList<Core.Src> srcs)
        {
            return FromSrc(srcs.Count, srcs);
        }

        public static StreamArray FromSrc(int numcells, IEnumerable<Core.Src> sidsrcs)
        {
            var num_shorts = numcells * 3;
            var array = new short[num_shorts];
            var stream = new StreamArray(array, StreamType.Src, numcells);

            int i = 0;
            int j = 0;
            foreach (var src in sidsrcs)
            {
                if (j >= numcells)
                {
                    break;
                }

                array[i++] = src.Section;
                array[i++] = src.Row;
                array[i++] = src.Cell;
                j++;
            }

            return stream;
        }
    }
}