using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public class ShapeSheetArrayBuilder<T>
    {
        protected List<T> items;
        private object[] object_array;

        public ShapeSheetArrayBuilder()
        {
            this.items = new List<T>();
        }

        public ShapeSheetArrayBuilder(int capacity)
        {
            this.items = new List<T>(capacity);
        }

        public int Count => this.items.Count;

        public void Add(T item)
        {
            this.items.Add(item);
            this.object_array = null;
        }

        public void AddRange(IEnumerable<T> items)
        {
            this.items.AddRange(items);
            this.object_array = null;
        }

        public void Clear()
        {
            this.items.Clear();
            this.object_array = null;
        }

        internal object[] ToObjectArray()
        {
            if (this.object_array != null)
            {
                return this.object_array;
            }

            this.object_array = new object[this.Count];
            for (int i = 0; i < this.Count; i++)
            {
                this.object_array[i] = this.items[i];
            }

            return this.object_array;
        }
    }
}