using System.Collections.Generic;

namespace VisioAutomation.Text.Markup
{
    public class NodeList<T> where T : Node
    {
        private readonly Node parent;
        private readonly List<T> items;

        public NodeList(Node parentnode)
        {
            this.parent = parentnode;
            this.items = new List<T>(0);
        }

        public IEnumerable<T> Items
        {
            get 
            {
                return this.items;
            }
        }

        public void Add(T item)
        {
            if (item.Parent != null)
            {
                if (item.Parent == this.parent)
                {
                    throw new System.ArgumentException("already a child of parent");
                }
                else
                {
                    throw new System.ArgumentException("already a child of another node");
                }
            }

            item.Parent = this.parent;
            this.items.Add(item);
        }

        public void Remove(T item)
        {
            if (item.Parent == null)
            {
                throw new System.ArgumentException("node does not have parent");
            }

            if (item.Parent != this.parent)
            {
                throw new System.ArgumentException("a child of a different parent");
            }

            this.items.Remove(item);
        }
        
        public int Count
        {
            get { return this.items.Count; }
        }
        
        public T this[int i]
        {
            get
            {
                return this.items[i];
            }
        }
    }
}