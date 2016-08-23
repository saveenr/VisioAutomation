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

        public void __Add_SIDSRC(short id, short sec, short row, short cell)
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

        public void __Add_SRC(short sec, short row, short cell)
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

    internal class StreamBuilderSRC: StreamBuilderBase
    {

        public StreamBuilderSRC(int capacity)
            : base(3, capacity)
        {
            
        }

        public void Add(short sec, short row, short cell)
        {
            this.__Add_SRC(sec, row, cell);
        }

        public void Add(SRC cell)
        {
            this.__Add_SRC(cell.Section, cell.Row, cell.Cell);
        }
    }

    internal class StreamBuilderSIDSRC : StreamBuilderBase
    {

        public StreamBuilderSIDSRC(int capacity) : base(4,capacity)
        {
        }

        public void Add(short shape_id, short sec, short row, short cell)
        {
            this.__Add_SIDSRC(shape_id, sec, row, cell);
        }

        public void Add(short shape_id, SRC cell)
        {
            this.__Add_SIDSRC(shape_id, cell.Section, cell.Row, cell.Cell);
        }
    }
}