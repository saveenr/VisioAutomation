namespace VisioAutomation.ShapeSheet.Query
{
    internal class StreamBuilder
    {
        public short[] Stream { get; }
        public int ChunksWrittenCount { get; private set; }
        public int ChunkSize { get; }
        public int ShortsWrittenCount { get; private set; }
        public int Capacity { get; }

        public StreamBuilder(int chunksize, int capacity)
        {
            if (chunksize != 3 && chunksize != 4)
            {
                string msg = "chunksize must be 3 or 4";
                throw new AutomationException(msg);                    
            }

            this.Capacity = capacity;
            this.Stream = new short[chunksize*capacity];
            this.ChunksWrittenCount = 0;
            this.ChunkSize = chunksize;
            this.ShortsWrittenCount = 0;
        }

        public void Add(short id, short sec, short row, short cell)
        {
            if (this.ChunkSize != 4)
            {
                string msg = "Only ChunkSize 4 supported";
                throw new AutomationException(msg);
            }

            if (this.ChunksWrittenCount >= this.Capacity)
            {
                string msg = "Exceeded Capacity";
                throw new AutomationException(msg);
            }

            this.Stream[this.ShortsWrittenCount++] = id;
            this.Stream[this.ShortsWrittenCount++] = sec;
            this.Stream[this.ShortsWrittenCount++] = row;
            this.Stream[this.ShortsWrittenCount++] = cell;
            this.ChunksWrittenCount++;
        }

        public void Add(short sec, short row, short cell)
        {
            if (this.ChunkSize != 3)
            {
                string msg = "Only ChunkSize 3 supported";
                throw new AutomationException(msg);
            }

            if (this.ChunksWrittenCount >= this.Capacity)
            {
                string msg = "Exceeded Capacity";
                throw new AutomationException(msg);
            }

            this.Stream[this.ShortsWrittenCount++] = sec;
            this.Stream[this.ShortsWrittenCount++] = row;
            this.Stream[this.ShortsWrittenCount++] = cell;
            this.ChunksWrittenCount++;
        }
    }
}