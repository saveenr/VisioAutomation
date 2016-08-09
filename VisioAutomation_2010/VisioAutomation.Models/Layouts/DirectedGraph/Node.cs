using System.Collections.Generic;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class Node
    {
        public string ID { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public string Label { get; set; }
        public DOM.BaseShape DOMNode { get; set; }
        public Dictionary<string, VACUSTPROP.CustomPropertyCells> CustomProperties { get; set; }

        public DOM.ShapeCells Cells;
    }
}