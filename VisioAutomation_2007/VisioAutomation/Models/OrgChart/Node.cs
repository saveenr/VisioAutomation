using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Models.OrgChart
{
    public class Node
    {
        private readonly NodeList _children;
        internal Node parent;

        public string Text { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public VA.DOM.Node DOMNode { get; set; }
        public string URL { get; set; }
        public VA.Drawing.Size? Size { get; set; }

        public Node()
        {
            this._children = new NodeList(this);
        }

        public Node(string name) :
            this ()
        {
            this.Text = name;
        }

        public NodeList Children
        {
            get { return _children; }
        }

        public Node Parent
        {
            get { return this.parent; }
        }      
    }
}