using System.Collections.Generic;

namespace VisioAutomation.Utilities
{
    public struct ArraySegment<T> : IEnumerable<T>
    {
        private readonly T[] Array;
        private readonly int _offset;
        private readonly int _count;

        public ArraySegment(T[] array, int offset, int count)
        {
            this.Array = array;
            this._offset = offset;
            this._count = count;
        }

        public T this[int index]
        {
            get
            {
                if (index >= this._count)
                {
                    throw new System.ArgumentOutOfRangeException(nameof(index));
                }

                var value = this.Array[this._offset + index];
               
                return value;
            }
        }

        public IEnumerator<T> GetEnumerator()
        {
            for (int i = 0; i < this._count; i++)
            {
                yield return this.Array[_offset + i];
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Count => this._count;

        public int Offset => this._offset;
    }
}