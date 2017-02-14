using System.Collections.Generic;
using Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public class UnitCodes
    {
        protected List<VisUnitCodes> items;
        private object[] object_array;

        public UnitCodes()
        {
            this.items = new List<VisUnitCodes>();
        }

        public UnitCodes(int capacity)
        {
            this.items = new List<VisUnitCodes>(capacity);
        }

        public int Count => this.items.Count;

        public void Add(Microsoft.Office.Interop.Visio.VisUnitCodes item)
        {
            this.items.Add(item);
            this.object_array = null;
        }

        public void AddRange(IEnumerable<VisUnitCodes> items)
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