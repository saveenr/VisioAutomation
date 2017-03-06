namespace VisioAutomation.Utilities
{
    internal struct SegmentedArray<T>
    {
        public readonly T[] Array;
        private int _capacity;
        private int _chucksize;

        public SegmentedArray(int capacity, int segmentsize)
        {
            int total_items = capacity * segmentsize;
            this._capacity = capacity;
            this._chucksize = segmentsize;
            this.Array = new T[total_items];
        }

        public VisioAutomation.Utilities.ArraySegment<T> this[int index]
        {
            get
            {
                int offset = (index * this._chucksize);
                return new VisioAutomation.Utilities.ArraySegment<T>(this.Array, offset, this._chucksize);
            }
        }

        public int Count => this._capacity;
    }
}