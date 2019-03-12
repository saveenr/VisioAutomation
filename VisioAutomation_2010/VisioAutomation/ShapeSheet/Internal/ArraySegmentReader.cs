namespace VisioAutomation.ShapeSheet.Internal
{
    public class ArraySegmentReader<T>
    {
        private readonly T[] array;
        private int pos;
        private int _count;

        public ArraySegmentReader(T[] array)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException(nameof(array));
            }
            this.array = array;
            this.pos = 0;
            this._count = 0;
        }

        public int Count => this._count;

        public int Capacity => this.array.Length;

        public VisioAutomation.ShapeSheet.Internal.ArraySegment<T> GetNextSegment(int size)
        {
            if (this.pos + size > this.array.Length)
            {
                throw new System.ArgumentOutOfRangeException(nameof(size));
            }
            var seg = new VisioAutomation.ShapeSheet.Internal.ArraySegment<T>(this.array, this.pos, size);
            this.pos += size;
            this._count += size;
            return seg;
        }
    }
}