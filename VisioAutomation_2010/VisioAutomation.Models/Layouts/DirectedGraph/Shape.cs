using VA = VisioAutomation;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class Shape : Node
    {
        public Shape(string id)
        {
            this.ID = id;
        }

        public string StencilName { get; set; }
        public string MasterName { get; set; }
        public string URL { get; set; }
        public VA.Drawing.Size? Size { get; set; }
        public System.Collections.Generic.List<DOM.Hyperlink> Hyperlinks { get; set; }
    }
}