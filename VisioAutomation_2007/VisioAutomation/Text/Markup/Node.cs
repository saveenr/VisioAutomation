using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Text.Markup
{
    public class Node
    {
        public NodeList<Node> Children { get; private set; }
        public NodeType NodeType { get; private set; }

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
            VA.Internal.TreeTraversal.EnumerateChildren<Node> enumchildren =
                n => (n.Children.Items);

            VA.Internal.TreeTraversal.EnterNode<Node> enternode =
                n => (n is TextElement);

            return VA.Internal.TreeTraversal.Walk<Node>(this, enumchildren, enternode);
        }

        public IEnumerable<Node> WalkNodes()
        {
            return VA.Internal.TreeTraversal.PreOrder<Node>(this,n=>n.Children.Items);
        }

    }
}
