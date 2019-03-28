using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class Node
    {
        private readonly NodeList _children;
        internal Node _parent;

        public string Text { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public Dom.Node DomNode { get; set; }
        public string Url { get; set; }
        public VisioAutomation.Geometry.Size? Size { get; set; }

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
            get { return this._children; }
        }

        public Node Parent
        {
            get { return this._parent; }
        }      
    }
}