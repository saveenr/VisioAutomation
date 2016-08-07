using System.Collections.Generic;

namespace VisioAutomation.Models.Text
{
    public class Node 
    {
        public NodeList<Node> Children { get;  set; }
        public NodeType NodeType { get; }

        internal Node(NodeType nt)
        {
            this.NodeType = nt;
            this.Children = new NodeList<Node>(this);
        }

        public string Name { get; protected set; }

        public Node Parent { get; internal set; }

        public string GetInnerText()
        {
            if (this.NodeType == NodeType.Literal)
            {
                var t = (Literal) this;
                return t.Text;
            }
            else if (this.NodeType == NodeType.Field)
            {
                var t = (Field)this;
                return t.PlaceholderText;
            }
            else if (this.NodeType == NodeType.Element)
            {
                var sb = new System.Text.StringBuilder();

                var entered_node_events = this.WalkNodes();

                foreach (var node in entered_node_events)
                {
                    if (node is Literal)
                    {
                        Literal mt = (Literal)node;
                        sb.Append(mt.Text);
                    }
                    else if (node is Field)
                    {
                        var t = (Field)node;
                        sb.Append(t.PlaceholderText);
                    }
                }

                return sb.ToString();
            }
            else
            {
                throw new System.InvalidOperationException();
            }
        }

        internal IEnumerable<Utilities.WalkEvent<Node>> Walk()
        {
            return Utilities.TreeOps.Walk<Node>(this, this.get_children_for_walk);
        }

        IEnumerable<Node> get_children_for_walk(Node n)
        {
            if (n is TextElement)
            {
                foreach (var c in n.Children)
                {
                    yield return c;
                }
            }
        }
        
        private IEnumerable<Node> WalkNodes()
        {
            return Utilities.TreeOps.PreOrder<Node>(this,n=>n.Children);
        }

        public void Add(Node n)
        {
            this.Children.Add(n);
        }
    }
}
