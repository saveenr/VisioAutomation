using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Models.Layouts.Tree
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
            get { return this._children; }
        }

        public VisioAutomation.Models.Text.Element Text { get; set;}
        public IVisio.Shape VisioShape { get; set; }
        public Dom.Node DOMNode { get; set; }
        public VA.Geometry.Size? Size { get; set; }
        public Dom.ShapeCells Cells { get; set; }

        public Node()
        {
            this._children = new NodeList(this);
        }

        public Node(string name)
            : this()
        {
            this.Text = new VisioAutomation.Models.Text.Element(name);
        }
    }
}