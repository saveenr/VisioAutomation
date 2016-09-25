namespace VisioAutomation.ShapeSheet.Queries.Utilities
{
    public class StreamBuilderBase
    {
        public short[] Stream { get; }
        public int Capacity { get; }

        private int _chunk_size { get; }
        private int _chunks_written_count { get; set; }
        private int _shorts_written_count { get; set; }

        public StreamBuilderBase(int chunksize, int capacity)
        {
            if (chunksize != 3 && chunksize != 4)
            {
                string msg = "chunksize must be 3 or 4";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            this.Capacity = capacity;
            this.Stream = new short[chunksize * capacity];
            this._chunks_written_count = 0;
            this._chunk_size = chunksize;
            this._shorts_written_count = 0;
        }

        protected void __Add_SIDSRC(short id, short sec, short row, short cell)
        {
            if (this._chunk_size != 4)
            {
                string msg = "Expected 3 shorts to add to stream";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            _check_enough_capacity();

            this.Stream[this._shorts_written_count++] = id;
            this.Stream[this._shorts_written_count++] = sec;
            this.Stream[this._shorts_written_count++] = row;
            this.Stream[this._shorts_written_count++] = cell;
            this._chunks_written_count++;
        }

        protected void __Add_SRC(short sec, short row, short cell)
        {
            if (this._chunk_size != 3)
            {
                string msg = "Expected 4 shorts to add to stream";
                throw new System.ArgumentOutOfRangeException(msg);
            }

            _check_enough_capacity();

            this.Stream[this._shorts_written_count++] = sec;
            this.Stream[this._shorts_written_count++] = row;
            this.Stream[this._shorts_written_count++] = cell;
            this._chunks_written_count++;
        }

        private void _check_enough_capacity()
        {
            if (this._chunks_written_count >= this.Capacity)
            {
                string msg = "Exceeded Capacity";
                throw new System.ArgumentOutOfRangeException(msg);
            }
        }

        public bool IsFull
        {
            get { return this._chunks_written_count == this.Capacity; }
        }

    }
}