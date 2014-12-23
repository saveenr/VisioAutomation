using System.Collections.Generic;
using System.Collections;

namespace VisioAutomation.Models.OrgChart
{
    public class NodeList : IEnumerable<Node>
    {
        private readonly Node parent;
        private readonly List<Node> items;

        public NodeList(Node parentnode)
        {
            this.parent = parentnode;
            this.items = new List<Node>(0);
        }

        public IEnumerator<Node> GetEnumerator()
        {
            foreach (var i in this.items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return GetEnumerator();
        }

        public void Add(Node item)
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

            item.parent = this.parent;
            this.items.Add(item);
        }

        public void Remove(Node item)
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
    }
}