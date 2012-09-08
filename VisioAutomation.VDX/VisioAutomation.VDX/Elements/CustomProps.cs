using System.Xml.Linq;
using System.Collections.Generic;
using VisioAutomation.VDX.Internal;
using VA=VisioAutomation;
using System.Linq;
using System.Collections;

namespace VisioAutomation.VDX.Elements
{
    public class CustomProps : IEnumerable<CustomProp>
    {
        private List<CustomProp> items;

        public CustomProps()
        {
            this.items = new List<CustomProp>();
        }

        public IEnumerator<CustomProp> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
        }

        public int Count
        {
            get { return this.items.Count;  }
        }

        public void Add( CustomProp cp)
        {
            cp.ID = this.items.Count + 1;
            this.items.Add(cp);
        }
    }
}