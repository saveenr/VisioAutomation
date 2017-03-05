using System.Windows.Forms.VisualStyles;

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

    }
}