using VisioAutomation.Shapes;


namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class Element
    {
        public string ID { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public string Label { get; set; }
        public Dom.BaseShape DomNode { get; set; }
        public CustomPropertyDictionary CustomProperties { get; set; }

        public Dom.ShapeCells Cells;
    }
}