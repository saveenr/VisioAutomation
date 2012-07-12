using VA = VisioAutomation;
using System.Collections.Generic;
using System.Collections;

namespace VisioAutomation.Text.Markup
{
    public class Node : IEnumerable<Node> 
    {
        private NodeList<Node> Children { get;  set; }
        public NodeType NodeType { get; private set; }

        internal Node(NodeType nt)
        {
            this.NodeType = nt;
            this.Children = new NodeList<Node>(this);
        }

        public IEnumerator<Node> GetEnumerator()
        {
            foreach (var i in this.Children.Items)
            {
                yield return i;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()     // Explicit implementation
        {                                           // keeps it hidden.
            return GetEnumerator();
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
            else if (this.NodeType == VA.Text.Markup.NodeType.Element)
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

        internal IEnumerable<VA.Internal.WalkEvent<Node>> Walk()
        {
            return VA.Internal.TreeTraversal.Walk<Node>(this, get_children_for_walk);
        }

        IEnumerable<Node> get_children_for_walk(Node n)
        {
            if (n is TextElement)
            {
                foreach (var c in n.Children.Items)
                {
                    yield return c;
                }
            }
        }
        
        private IEnumerable<Node> WalkNodes()
        {
            return VA.Internal.TreeTraversal.PreOrder<Node>(this,n=>n);
        }

        public void Add(Node n)
        {
            this.Children.Add(n);
        }
    }
}
