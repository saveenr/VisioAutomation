namespace VisioAutomation.ShapeSheet.Streams
{
    public struct StreamArray
    {
        public readonly short[] Array;
        internal Internal.CoordType CoordType;
        public readonly int ChunkSize;

        internal StreamArray(short[] array, Internal.CoordType coord)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException(nameof(array));
            }

            this.Array = array;
            this.CoordType = coord;
            this.ChunkSize = coord == Internal.CoordType.SidSrc ? 4 : 3;
        }

    }
}