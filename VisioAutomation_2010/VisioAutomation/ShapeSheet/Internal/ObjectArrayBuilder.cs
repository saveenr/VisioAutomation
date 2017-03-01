using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet.Internal
{
    internal class ObjectArrayBuilder<T>
    {
        protected List<T> items;

        public ObjectArrayBuilder()
        {
            this.items = new List<T>();
        }

        public ObjectArrayBuilder(int capacity)
        {
            this.items = new List<T>(capacity);
        }

        public int Count => this.items.Count;

        public void Add(T item)
        {
            this.items.Add(item);
        }

        public void AddRange(IEnumerable<T> items)
        {
            this.items.AddRange(items);
        }

        public void Clear()
        {
            this.items.Clear();
        }

        public object[] ToObjectArray()
        {
            var object_array = new object[this.Count];
            for (int i = 0; i < this.Count; i++)
            {
                object_array[i] = this.items[i];
            }

            return object_array;
        }
    }
}