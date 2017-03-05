namespace VisioAutomation.ShapeSheet.Streams
{
    public struct StreamArray
    {
        public readonly short[] Array;
        public readonly Streams.StreamType StreamType;
        public readonly int ChunkSize;

        internal StreamArray(short[] array, Streams.StreamType cell_coord)
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
        }

    }
}