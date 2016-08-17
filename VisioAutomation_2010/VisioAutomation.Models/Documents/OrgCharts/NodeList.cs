using System.Collections.Generic;
using System.Collections;

namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class NodeList : IEnumerable<Node>
    {
        private readonly Node _parent;
        private readonly List<Node> _items;

        public NodeList(Node parentnode)
        {
            this._parent = parentnode;
            this._items = new List<Node>(0);
        }

        public IEnumerator<Node> GetEnumerator()
        {
            foreach (var i in this._items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return this.GetEnumerator();
        }

        public void Add(Node item)
        {
            if (item.Parent != null)
            {
                if (item.Parent == this._parent)
                {
                    throw new System.ArgumentException("already a child of parent");
                }
                else
                {
                    throw new System.ArgumentException("already a child of another node");
                }
            }

            item._parent = this._parent;
            this._items.Add(item);
        }

        public void Remove(Node item)
        {
            if (item.Parent == null)
            {
                throw new System.ArgumentException("node does not have parent");
            }

            if (item.Parent != this._parent)
            {
                throw new System.ArgumentException("already a child of a different parent");
            }

            this._items.Remove(item);
        }

        public int Count
        {
            get { return this._items.Count; }
        }
    }
}