using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.BoxLayout
{
    public class Node
    {
        private List<Node> m_children;
        internal Node parent;

        public double? Width { get; set; }
        public double? Height { get; set; }
        public double Padding { get; set; }
        public double ChildSeparation { get; set; }

        public object Data { get; set; }
        public VA.Drawing.Rectangle Rectangle { get; set; }
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

        public Node Parent
        {
            get { return this.parent; }
        }

        public IEnumerable<Node> Children
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

        public Node AddNode(double? width, double? height)
        {
            var node = new Node();
            node.Width = width;
            node.Height = height;

            this.AddNode(node);

            return node;
        }

        public Node AddRow()
        {
            return this.AddNode(LayoutDirection.Horizonal);
        }

        public Node AddRow(double? width, double? height, VA.Drawing.AlignmentVertical valign)
        {
            var node = new Node();
            node.Width = width;
            node.Height = height;
            node.AlignmentVertical = valign;

            this.AddNode(node);

            return node;
        }

        public Node AddColumn()
        {
            return this.AddNode(LayoutDirection.Vertical);
        }

        public Node AddColumn(double? width, double? height, VA.Drawing.AlignmentHorizontal halign)
        {
            var node = new Node();
            node.Width = width;
            node.Height = height;
            node.AlignmentHorizontal = halign;

            this.AddNode(node);

            return node;
        }

        public Node AddNode(Node node)
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
                this.m_children = new List<Node>();
            }
            this.m_children.Add(node);
            node.parent = this;

            return node;
        }

        private Node AddNode(LayoutDirection dir)
        {
            var node = new Node();
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