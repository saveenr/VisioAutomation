namespace VisioAutomation.Collections
{
    public class ArraySegmentReader<T>
    {
        private readonly T[] _array;
        private int _pos;
        private int _count;

        public ArraySegmentReader(T[] array)
        {
            if (array == null)
            {
                throw new System.ArgumentNullException(nameof(array));
            }
            this._array = array;
            this._pos = 0;
            this._count = 0;
        }

        public int Count => this._count;

        public int Capacity => this._array.Length;

        public VisioAutomation.Collections.ArraySegment<T> GetNextSegment(int size)
        {
            if (this._pos + size > this._array.Length)
            {
                throw new System.ArgumentOutOfRangeException(nameof(size));
            }
            var seg = new VisioAutomation.Collections.ArraySegment<T>(this._array, this._pos, size);
            this._pos += size;
            this._count += size;
            return seg;
        }
    }
}