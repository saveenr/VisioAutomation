using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public class Formulas
    {
        protected List<string> items;
        private object[] object_array;

        public Formulas()
        {
            this.items = new List<string>();
        }

        public Formulas(int capacity)
        {
            this.items = new List<string>(capacity);
        }

        public int Count => this.items.Count;

        public void Add(string item)
        {
            this.items.Add(item);
            this.object_array = null;
        }

        public void AddRange(IEnumerable<string> items)
        {
            this.items.AddRange(items);
            this.object_array = null;
        }

        public void Clear()
        {
            this.items.Clear();
            this.object_array = null;
        }

        public object[] ToObjectArray()
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