using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Internal
{
    public struct ArraySegment<T> : IEnumerable<T>
    {
        private System.ArraySegment<T> sas;
        
        public ArraySegment(T[] array, int offset, int count)
        {
            this.sas = new System.ArraySegment<T>(array, offset, count);
        }

        public T this[int index]
        {
            get { return get_value_at_index(index); }
            set { set_value_at_index(index, value); }
        }

        private void set_value_at_index(int index, T value)
        {
            validate_index(index);
            this.sas.Array[this.Offset + index] = value;
        }

        private T get_value_at_index(int index)
        {
            validate_index(index);
            var value = this.sas.Array[this.Offset + index];
            return value;
        }

        private void validate_index(int index)
        {
            if ((index < 0) && (index >= this.sas.Count))
            {
                throw new System.ArgumentOutOfRangeException(nameof(index));
            }
        }

        public IEnumerator<T> GetEnumerator()
        {
            for (int i = 0; i < this.sas.Count; i++)
            {
                yield return this.sas.Array[this.sas.Offset + i];
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public int Count => this.sas.Count;

        public int Offset => this.sas.Offset;
    }
}