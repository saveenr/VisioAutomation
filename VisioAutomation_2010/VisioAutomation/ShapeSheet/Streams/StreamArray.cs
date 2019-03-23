using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Streams
{
    public struct StreamArray
    {
        public readonly short[] Array;
        public readonly Streams.StreamType StreamType;
        public readonly int ChunkSize;
        public readonly int Count;

        internal StreamArray(short[] array, Streams.StreamType cell_coord, int count)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException(nameof(array));
            }

            this.Array = array;
            this.StreamType = cell_coord;
            this.ChunkSize = cell_coord == Streams.StreamType.SidSrc ? 4 : 3;

            if (array.Length % this.ChunkSize != 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Coordinate type and length of array to not match");
            }

            if (count * this.ChunkSize != array.Length)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Count does not match the number of short elements in the array");
            }

            this.Count = count;
        }



        public static VisioAutomation.ShapeSheet.Streams.StreamArray CreateSidSrcStream(IList<SidSrc> sidsrcs)
        {
            return CreateSidSrcStream(sidsrcs.Count, sidsrcs);
        }

        public static VisioAutomation.ShapeSheet.Streams.StreamArray CreateSidSrcStream(int numcells, IEnumerable<SidSrc> sidsrcs)
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
        public static VisioAutomation.ShapeSheet.Streams.StreamArray CreateSrcStream(IList<Src> srcs)
        {
            return CreateSrcStream(srcs.Count, srcs);
        }
        public static VisioAutomation.ShapeSheet.Streams.StreamArray CreateSrcStream(int numcells, IEnumerable<Src> sidsrcs)
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