using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.Streams;

namespace VisioAutomation.ShapeSheet.Internal
{
    class WriterCollection_Src
    {
        private List<SrcWrite> items;

        public WriterCollection_Src()
        {
            this.items = new List<SrcWrite>();
        }

        public void Clear()
        {
            this.items.Clear();
        }

        public void Add(Src src, string value)
        {
            var item = new SrcWrite(src, value);
            this.items.Add(item);
        }

        public Streams.StreamArray BuildStream()
        {
            var streambuilder = new FixedSrcStreamBuilder(this.items.Count);
            streambuilder.AddRange(this.items.Select(i => i.Src));
            return streambuilder.ToStream();
        }

        public object[] BuildValues()
        {
            var array = new object[this.items.Count];
            for (int i = 0; i < this.items.Count; i++)
            {
                array[i] = this.items[i].Value;
            }
            return array;
        }

        public int Count => this.items.Count;

        struct SrcWrite
        {
            public Src Src;
            public string Value;

            public SrcWrite(Src src, string value)
            {
                this.Src = src;
                this.Value = value;
            }
        }
    }
}