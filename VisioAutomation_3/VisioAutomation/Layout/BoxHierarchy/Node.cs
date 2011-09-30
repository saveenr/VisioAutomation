using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.BoxHierarchy
{
    public class Node<T>
    {
        private List<Node<T>> m_children;
        internal Node<T> parent;

        public double? Width { get; set; }
        public double? Height { get; set; }
        public double Padding { get; set; }
        public double ChildSeparation { get; set; }

        public T Data { get; set; }
        public VA.Drawing.Rectangle Rectangle { get; set; }
        public VA.Drawing.Rectangle ReservedRectangle { get; set; }
        public LayoutDirection Direction;
        public VA.Drawing.AlignmentHorizontal AlignmentHorizontal { get; set; }
        public VA.Drawing.AlignmentVertical AlignmentVertical { get; set; }

        public Node() : 
            this ( LayoutDirection.Vertical)
        {
        }

        public Node(LayoutDirection dir)
        {
            this.Direction = dir;
            this.m_children = null;
        }

        public Node<T> Parent
        {
            get { return this.parent; }
        }

        public IEnumerable<Node<T>> Children
        {
            get
            {
                if (this.m_children == null)
                {
                    yield break;
                }
                else
                {
                    foreach (var c in this.m_children)
                    {
                        yield return c;
                    }
                }
            }
        }

        public Node<T> AddNode(double? width, double? height)
        {
            var node = new Node<T>();
            node.Width = width;
            node.Height = height;

            this.AddNode(node);

            return node;
        }

        public Node<T> AddNode(double? width, double? height, VA.Drawing.AlignmentHorizontal halign)
        {
            var node = new Node<T>();
            node.Width = width;
            node.Height = height;
            node.AlignmentHorizontal = halign;

            this.AddNode(node);

            return node;
        }

        public Node<T> AddNode(double? width, double? height, VA.Drawing.AlignmentVertical valign)
        {
            var node = new Node<T>();
            node.Width = width;
            node.Height = height;
            node.AlignmentVertical = valign;

            this.AddNode(node);

            return node;
        }

        public Node<T> AddNode(Node<T> node)
        {
            if (node == null)
            {
                throw new System.ArgumentNullException("node");
            }

            if (node.parent == this)
            {
                throw new System.ArgumentException("node already a child of this this node");
            }

            if (node.parent != null)
            {
                throw new System.ArgumentException("node already has a parent");
            }

            if (this.m_children == null)
            {
                this.m_children = new List<Node<T>>();
            }
            this.m_children.Add(node);
            node.parent = this;

            return node;
        }

        public Node<T> AddNode(LayoutDirection dir)
        {
            var node = new Node<T>();
            node.Direction = dir;

            return this.AddNode(node);
        }

        public int ChildCount
        {
            get
            {
                if (this.m_children == null)
                {
                    return 0;
                }
                else
                {
                    return this.m_children.Count;
                }
            }
        }
    }
}