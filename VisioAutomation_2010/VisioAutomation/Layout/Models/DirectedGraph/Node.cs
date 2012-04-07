using System.Collections.Generic;
using VA=VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.Models.DirectedGraph
{
    public class Node
    {
        public string ID { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public string Label { get; set; }
        public VA.DOM.Shape DOMNode { get; set; }
        public Dictionary<string, VA.CustomProperties.CustomPropertyCells> CustomProperties { get; set; }

        public VA.DOM.ShapeCells Cells;
    }
}