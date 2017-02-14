using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeSheet
{
    public abstract class Stream
    {
        public abstract short[] ToStreamArray();
        public abstract int Count();

        public virtual void AddSIDSRC(SIDSRC sidsrc)
        {
            throw new System.NotImplementedException();
        }

        public virtual void AddSRC(SRC src)
        {
            throw new System.NotImplementedException();
        }
    }

    public abstract class Stream<T> : Stream
    {
        protected List<T> items;
        private short[] stream;

        public Stream()
        {
            this.items = new List<T>();
        }

        public override int Count()
        {
            return this.items.Count;
        }

        public Stream(int capacity)
        {
            this.items = new List<T>(capacity);
        }

        public void Add(T item)
        {
            this.items.Add(item);
            this.stream = null;
        }

        public void AddRange(IEnumerable<T> items)
        {
            this.items.AddRange(items);
            this.stream = null;
        }

        public void Clear()
        {
            this.items.Clear();
            this.stream = null;
        }

        public override short[] ToStreamArray()
        {
            if (this.stream != null)
            {
                return this.stream;
            }
            else
            {
                this.stream = this.get_stream();
                return this.stream;
            }
        }

        protected abstract short[] get_stream();
    }

    public class UnitCodes
    {
        protected List<IVisio.VisUnitCodes> items;
        private object[] object_array;

        public UnitCodes()
        {
            this.items = new List<IVisio.VisUnitCodes>();
        }

        public UnitCodes(int capacity)
        {
            this.items = new List<IVisio.VisUnitCodes>(capacity);
        }

        public int Count => this.items.Count;

        public void Add(IVisio.VisUnitCodes item)
        {
            this.items.Add(item);
            this.object_array = null;
        }

        public void AddRange(IEnumerable<IVisio.VisUnitCodes> items)
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
