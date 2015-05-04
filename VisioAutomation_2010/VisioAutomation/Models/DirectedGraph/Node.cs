using System.Collections.Generic;
using CUSTPROP=VisioAutomation.Shapes.CustomProperties;
using VA=VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.DirectedGraph
{
    public class Node
    {
        public string ID { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public string Label { get; set; }
        public DOM.BaseShape DOMNode { get; set; }
        public Dictionary<string, CUSTPROP.CustomPropertyCells> CustomProperties { get; set; }

        public DOM.ShapeCells Cells;
    }
}