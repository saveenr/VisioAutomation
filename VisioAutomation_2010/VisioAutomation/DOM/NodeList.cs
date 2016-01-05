using System.Collections;
using System.Collections.Generic;

namespace VisioAutomation.DOM
{
    public class NodeList<T> : IEnumerable<T> where T : Node
    {
        private List<T> items;

        public Node Parent { get; }

        internal NodeList(Node parent)
        {
            this.Parent = parent;
            this.items = null;
        }

        public IEnumerator<T> GetEnumerator()
        {
            foreach (var i in this.GetItems())
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     
        {                                           
            return this.GetEnumerator();
        }

        private IEnumerable<T> GetItems()
        {
            if (this.items == null)
            {
                yield break;
            }

            foreach (T n in this.items)
            {
                yield return n;
            }
        }

        /// <summary>
        /// Adds a node as a child
        /// </summary>
        /// <param name="node_to_add"></param>
        /// <returns></returns>
        public T Add(T node_to_add)
        {
            if (node_to_add == null)
            {
                throw new System.ArgumentNullException(nameof(node_to_add));
            }

            if (node_to_add == this.Parent)
            {
                throw new System.ArgumentException("Cannot add node as a child of itself");
            }

            if (node_to_add.Parent != null)
            {
                throw new System.ArgumentException("Node is already a child of a node");
            }

            this.items = this.items ?? new List<T>();
            node_to_add.Parent = this.Parent;
            this.items.Add(node_to_add);

            return node_to_add;
        }

        /// <summary>
        /// Adds a set of nodes as children
        /// </summary>
        /// <param name="nodes"></param>
        public void Add(IEnumerable<T> nodes)
        {
            if (nodes == null)
            {
                throw new System.ArgumentNullException(nameof(nodes));
            }

            foreach (T i in nodes)
            {
                this.Add(i);
            }
        }

        public int Count
        {
            get
            {
                if (this.items == null)
                {
                    return 0;
                }

                return this.items.Count;
            }
        }

        public T this[int index]
        {
            get { return this.items[index]; }
        }
    }
}