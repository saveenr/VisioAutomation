using System.Collections.Generic;

namespace VisioAutomation.Utilities
{
    public struct ArraySegment<T> : IEnumerable<T>
    {
        private readonly T[] Array;
        private readonly int _offset;
        private readonly int _length;

        public ArraySegment(T[] array, int offset, int length)
        {
            this.Array = array;
            this._offset = offset;
            this._length = length;
        }

        public T this[int index]
        {
            get { return set_value_at_index(index); }
        }

        private T set_value_at_index(int index)
        {
            validate_index(index);

            var value = this.Array[this._offset + index];
            return value;
        }

        private void validate_index(int index)
        {
            if ((index < 0) && (index >= this._length))
            {
                throw new System.ArgumentOutOfRangeException(nameof(index));
            }
        }

        public IEnumerator<T> GetEnumerator()
        {
            for (int i = 0; i < this._length; i++)
            {
                yield return this.Array[_offset + i];
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Length => this._length;

        public int Offset => this._offset;
    }
}