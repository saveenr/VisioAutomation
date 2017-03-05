namespace VisioAutomation.ShapeSheet.Streams
{
    public struct StreamArray
    {
        public readonly short[] Array;
        internal Internal.CellCoord CellCoord;
        public readonly int ChunkSize;

        internal StreamArray(short[] array, Internal.CellCoord cell_coord)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException(nameof(array));
            }

            this.Array = array;
            this.CellCoord = cell_coord;
            this.ChunkSize = cell_coord == Internal.CellCoord.SidSrc ? 4 : 3;

            if (array.Length % this.ChunkSize != 0)
            {
                throw new VisioAutomation.Exceptions.InternalAssertionException("Coordinate type and length of array to not match");
            }
        }

    }
}