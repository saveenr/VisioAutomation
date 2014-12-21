using System.Collections.Generic;
using VisioAutomation.Shapes.CustomProperties;
using VA=VisioAutomation;
using IVisio=Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.DirectedGraph
{
    public class Node
    {
        public string ID { get; set; }
        public IVisio.Shape VisioShape { get; set; }
        public string Label { get; set; }
        public VA.DOM.BaseShape DOMNode { get; set; }
        public Dictionary<string, CustomPropertyCells> CustomProperties { get; set; }

        public VA.DOM.ShapeCells Cells;
    }
}