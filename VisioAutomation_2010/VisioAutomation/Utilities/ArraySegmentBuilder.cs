namespace VisioAutomation.Utilities
{
    internal class ArraySegmentBuilder<T>
    {
        private T[] array;
        private int pos;
        private int _count;

        public ArraySegmentBuilder(T[] array)
        {
            this.array = array;
            this.pos = 0;
            this._count = 0;
        }

        public int Count => _count;

        public VisioAutomation.Utilities.ArraySegment<T> GetNextSegment(int size)
        {
            var seg = new VisioAutomation.Utilities.ArraySegment<T>(this.array, this.pos, size);
            this.pos += size;
            this._count += size;
            return seg;
        }
    }
}