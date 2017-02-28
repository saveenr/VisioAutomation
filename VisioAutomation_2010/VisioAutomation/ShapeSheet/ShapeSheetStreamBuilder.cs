using System.Collections.Generic;

namespace VisioAutomation.ShapeSheet
{
    public abstract class ShapeSheetStreamBuilder
    {
        public abstract short[] ToStream();

        public abstract int Count();
    }

    public abstract class ShapeSheetStreamBuilder<T> : ShapeSheetStreamBuilder
    {
        protected List<T> items;

        public ShapeSheetStreamBuilder()
        {
            this.items = new List<T>();
        }

        public override int Count() => this.items.Count;

        public ShapeSheetStreamBuilder(int capacity)
        {
            this.items = new List<T>(capacity);
        }

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

        public override short[] ToStream()
        {
             var stream = this.build_stream();
             return stream;
        }

        protected abstract short[] build_stream();
    }
}



