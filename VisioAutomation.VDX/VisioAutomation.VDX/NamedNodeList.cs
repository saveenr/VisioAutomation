using System.Collections.Generic;
using System.Linq;

namespace VisioAutomation.VDX
{
    public class NamedNodeList<T> : Elements.Node where T : Elements.Node
    {
        private readonly Dictionary<string, T> dic;
        private readonly List<T> items;
        private readonly System.Func<T, string> func_get_name;

        public NamedNodeList(System.Func<T, string> func_get_name)
        {
            if (func_get_name == null)
            {
                throw new System.ArgumentNullException("func_get_name");
            }

            this.items = new List<T>();
            var compare = System.StringComparer.InvariantCultureIgnoreCase;
            this.dic = new Dictionary<string, T>(compare);
            this.func_get_name = func_get_name;
        }

        public bool ContainsName(string name)
        {
            if (name == null)
            {
                throw new System.ArgumentNullException("name");
            }

            return this.dic.ContainsKey(name);
        }

        public bool Contains(T item)
        {
            if (item == null)
            {
                throw new System.ArgumentNullException("item");
            }

            return (item._parent == this);
        }

        public virtual void Add(T item)
        {
            if (item == null)
            {
                throw new System.ArgumentNullException("item");
            }

            if (item._parent == this)
            {
                //throw new System.ArgumentException("item is already a member of this collection");
            }

            if (item._parent != null)
            {
                throw new System.ArgumentException("item is already a member of another collection");
            }

            item._parent = this;
            string name = this.func_get_name(item);

            if (this.ContainsName(name))
            {
                throw new System.ArgumentException("Already contains an item with that name", "item");
            }
            else
            {
                this.dic[name] = item;
                this.items.Add(item);
            }
        }

        public T this[string name]
        {
            get { return this.dic[name]; }
        }

        public int Count
        {
            get { return this.items.Count; }
        }

        public IEnumerable<T> Items
        {
            get { return this.items.AsEnumerable(); }
        }
    }
}