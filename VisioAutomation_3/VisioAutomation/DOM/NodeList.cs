using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class NodeList<T> where T : Node
    {
        private List<T> items;

        public Node Parent { get; private set; }

        internal NodeList(Node parent)
        {
            this.Parent = parent;
            this.items = null;
        }

        /// <summary>
        /// Enumerates through all the child nodes
        /// </summary>
        /// <returns></returns>
        public IEnumerable<T> Items
        {
            get
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
                throw new System.ArgumentNullException("node_to_add");
            }

            if (node_to_add == this.Parent)
            {
                throw new System.ArgumentException("Cannot add node as a child of itself");
            }

            if (node_to_add.Parent != null)
            {
                throw new System.ArgumentException("already a child of a node");
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
                throw new System.ArgumentNullException("nodes");
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