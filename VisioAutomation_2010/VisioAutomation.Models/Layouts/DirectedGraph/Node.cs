using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class Node
    {
        public string ID { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public string Label { get; set; }
        public Dom.BaseShape DOMNode { get; set; }
        public VisioAutomation.Shapes.CustomProperties.CustomPropertyDictionary CustomProperties { get; set; }

        public Dom.ShapeCells Cells;
    }
}