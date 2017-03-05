using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.Streams;

namespace VisioAutomation.ShapeSheet.Internal
{
    class WriterCollection_SidSrc
    {
        private List<SidSrcWrite> items;

        public WriterCollection_SidSrc()
        {
            this.items = new List<SidSrcWrite>();
        }

        public void Clear()
        {
            this.items.Clear();
        }

        public void Add(SidSrc sidsrc, string value)
        {
            var item = new SidSrcWrite(sidsrc,value);
            this.items.Add(item);
        }

        public Streams.StreamArray BuildStream()
        {
            var streambuilder = new FixedSidSrcStreamBuilder(this.items.Count);
            streambuilder.AddRange(this.items.Select( i=>i.SidSrc));
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

        struct SidSrcWrite
        {
            public SidSrc SidSrc;
            public string Value;

            public SidSrcWrite(SidSrc sidsrc, string value)
            {
                this.SidSrc = sidsrc;
                this.Value = value;
            }
        }
    }
}