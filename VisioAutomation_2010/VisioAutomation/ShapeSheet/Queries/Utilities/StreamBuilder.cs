namespace VisioAutomation.ShapeSheet.Queries.Utilities
{
    class StreamBuilder
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
                throw new System.ArgumentOutOfRangeException(msg);
            }

            this.Capacity = capacity;
            this.Stream = new short[chunksize * capacity];
            this.ChunksWrittenCount = 0;
            this.ChunkSize = chunksize;
            this.ShortsWrittenCount = 0;
        }

        public void Add(short id, short sec, short row, short cell)
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

        public void Add(short sec, short row, short cell)
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
    }

    internal class StreamBuilderSRC
    {
        private StreamBuilder builder;

        public int Capacity
        {
            get { return this.builder.Capacity;  }
        }

        public StreamBuilderSRC(int capacity)
        {
            this.builder = new StreamBuilder(3, capacity);
        }

        public void Add(short sec, short row, short cell)
        {
            this.builder.Add(sec, row, cell);
        }

        public void Add(SRC cell)
        {
            this.builder.Add(cell.Section, cell.Row, cell.Cell);
        }

        public short[] Stream
        {
            get { return this.builder.Stream; }
        }

        public bool IsFull
        {
            get { return this.builder.ChunksWrittenCount == this.Capacity; }
        }
    }

    internal class StreamBuilderSIDSRC
    {
        private StreamBuilder builder;

        public int Capacity
        {
            get { return this.builder.Capacity; }
        }

        public StreamBuilderSIDSRC(int capacity)
        {
            this.builder = new StreamBuilder(4, capacity);
        }

        public void Add(short shape_id, short sec, short row, short cell)
        {
            this.builder.Add(shape_id, sec, row, cell);
        }

        public void Add(short shape_id, SRC cell)
        {
            this.builder.Add(shape_id, cell.Section, cell.Row, cell.Cell);
        }

        public short[] Stream
        {
            get { return this.builder.Stream; }
        }

        public bool IsFull
        {
            get { return this.builder.ChunksWrittenCount == this.Capacity; }
        }

    }

}