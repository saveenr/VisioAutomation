namespace VisioAutomation.ShapeSheet.Queries.Utilities
{
    class StreamBuilderBase
    {
        public short[] Stream { get; }
        private int ChunksWrittenCount { get; set; }
        private int ChunkSize { get; }
        private int ShortsWrittenCount { get; set; }
        public int Capacity { get; }

        public StreamBuilderBase(int chunksize, int capacity)
        {
            if (chunksize != 3 && chunksize != 4)
            {
                string msg = "chunksize must be 3 or 4";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            this.Capacity = capacity;
            this.Stream = new short[chunksize * capacity];
            this.ChunksWrittenCount = 0;
            this.ChunkSize = chunksize;
            this.ShortsWrittenCount = 0;
        }

        protected void __Add_SIDSRC(short id, short sec, short row, short cell)
        {
            if (this.ChunkSize != 4)
            {
                string msg = "Only ChunkSize 4 supported";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            if (this.ChunksWrittenCount >= this.Capacity)
            {
                string msg = "Exceeded Capacity";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            this.Stream[this.ShortsWrittenCount++] = id;
            this.Stream[this.ShortsWrittenCount++] = sec;
            this.Stream[this.ShortsWrittenCount++] = row;
            this.Stream[this.ShortsWrittenCount++] = cell;
            this.ChunksWrittenCount++;
        }

        protected void __Add_SRC(short sec, short row, short cell)
        {
            if (this.ChunkSize != 3)
            {
                string msg = "Only ChunkSize 3 supported";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            if (this.ChunksWrittenCount >= this.Capacity)
            {
                string msg = "Exceeded Capacity";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            this.Stream[this.ShortsWrittenCount++] = sec;
            this.Stream[this.ShortsWrittenCount++] = row;
            this.Stream[this.ShortsWrittenCount++] = cell;
            this.ChunksWrittenCount++;
        }

        public bool IsFull
        {
            get { return this.ChunksWrittenCount == this.Capacity; }
        }

    }
}