[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("VTest")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("VTest.Scripting")]
namespace VisioAutomation.Internal
{
    /// <summary>
    /// Allows the incremental building up of segments from an array
    /// </summary>
    /// <typeparam name="T"></typeparam>
    ///
    ///
    internal class ArraySegmentEnumerator<T>
    {
        private readonly T[] _array;
        private int _pos;
        private int _count;

        public ArraySegmentEnumerator(T[] array)
        {
            this._array = array ?? throw new System.ArgumentNullException(nameof(array));
            this._pos = 0;
            this._count = 0;
        }

        public int Count => this._count;

        public int Capacity => this._array.Length;

        public ArraySegment<T> GetNextSegment(int size)
        {
            // Keep in mind its ALWAYS OK to ask for a size of zero
            // even if there's nothing left to enumerte

            if (size < 0)
            {
                // there's nothing left to consume
                string msg = string.Format("Size must be positive. Actual value given is {0}", size);
                throw new System.ArgumentOutOfRangeException(nameof(size), msg);
            }

            if (size >0 && this.Count == this.Capacity)
            {
                // there's nothing left to consume
                string msg = string.Format("All {0} elements of the array have been consumed", this._array.Length);
                throw new System.ArgumentOutOfRangeException(nameof(size), msg);
            }

            if (this._pos + size > this._array.Length)
            {
                int remaining = this.Capacity - this.Count;
                // there's request goes beyond want is available
                string msg = string.Format("Requesting {0} elements but only {1} are remaining", size, remaining);
                throw new System.ArgumentOutOfRangeException(nameof(size),msg);
            }

            var seg = new ArraySegment<T>(this._array, this._pos, size);
            this._pos += size;
            this._count += size;
            return seg;
        }
    }
}