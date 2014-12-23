using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Models.Tree
{
    public class Node
    {
        private readonly NodeList _children;
        internal Node parent;

        public Node Parent
        {
            get { return this.parent; }
        }

        public NodeList Children
        {
            get { return _children; }
        }

        public VA.Text.Markup.TextElement Text { get; set;}
        public IVisio.Shape VisioShape { get; set; }
        public VA.DOM.Node DOMNode { get; set; }
        public VA.Drawing.Size? Size { get; set; }
        public VA.DOM.ShapeCells Cells { get; set; }

        public Node()
        {
            this._children = new NodeList(this);
        }

        public Node(string name)
            : this()
        {
            this.Text = new VA.Text.Markup.TextElement(name);
        }
    }
}