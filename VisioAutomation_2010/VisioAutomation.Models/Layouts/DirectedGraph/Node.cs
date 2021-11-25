using VA = VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class Node : Element
    {
        public Node(string id)
        {
            this.ID = id;
        }

        public string StencilName { get; set; }
        public string MasterName { get; set; }
        public string Url { get; set; }
        public VA.Core.Size? Size { get; set; }
        public System.Collections.Generic.List<Dom.Hyperlink> Hyperlinks { get; set; }
    }
}