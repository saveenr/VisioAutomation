using System.Collections.Generic;
using VA = VisioAutomation;

namespace VisioAutomation.Layout.BoxLayout
{
    enum NodeType
    {
         Box,
         Container
    }

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

        private NodeType nodetype;

        private Node(NodeType type, LayoutDirection dir)
        {
            this.nodetype = type;
            this.Direction = dir;
            this.m_children = null;
        }

        internal static Node NewContainer(LayoutDirection dir)
        {
            return new Node(NodeType.Container, dir);
        }

        internal static Node NewBox()
        {
            return new Node(NodeType.Box,LayoutDirection.Vertical);
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

        public Node AddBox(double? width, double? height)
        {
            if (this.nodetype == NodeType.Box)
            {
                throw new AutomationException("Can't add Boxes to Boxes");
            }

            var node = Node.NewBox();
            node.Width = width;
            node.Height = height;

            this._addnode(node);

            return node;
        }

        public Node AddRow()
        {
            if (this.nodetype == NodeType.Box)
            {
                throw new AutomationException("Can't add Rows to Boxes");
            }

            var n = Node.NewContainer(LayoutDirection.Horizontal);
            this._addnode(n);
            return n;
        }

        public Node AddRow(double? width, double? height, VA.Drawing.AlignmentVertical valign)
        {
            if (this.nodetype == NodeType.Box)
            {
                throw new AutomationException("Can't add Rows to Boxes");
            }

            var node = Node.NewBox();
            node.Width = width;
            node.Height = height;
            node.AlignmentVertical = valign;

            this._addnode(node);

            return node;
        }

        public Node AddColumn()
        {
            if (this.nodetype == NodeType.Box)
            {
                throw new AutomationException("Can't add Columns to Boxes");
            }

            var n = Node.NewContainer(LayoutDirection.Vertical);
            this._addnode(n);
            return n;
        }

        public Node AddColumn(double? width, double? height, VA.Drawing.AlignmentHorizontal halign)
        {
            if (this.nodetype == NodeType.Box)
            {
                throw new AutomationException("Can't add Columns to Boxes");
            }

            var node = Node.NewContainer(LayoutDirection.Vertical);
            node.Width = width;
            node.Height = height;
            node.AlignmentHorizontal = halign;

            this._addnode(node);

            return node;
        }

        private Node _addnode(Node node)
        {
            if (this.nodetype == NodeType.Box)
            {
                throw new AutomationException("Can't add nodes to Boxes");
            }

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