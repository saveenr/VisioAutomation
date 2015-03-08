using System.Collections.Generic;
using CUSTPROP=VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.DOM
{
    public class BaseShape : Node
    {
        public IVisio.Shape VisioShape { get; set; }
        public short VisioShapeID { get; internal set; }

        public VA.Text.Markup.TextElement Text { get; set; }
        public Dictionary<string, CUSTPROP.CustomPropertyCells> CustomProperties { get; set; }
        public List<Hyperlink> Hyperlinks { get; set; }

        // Be aware that if multiple nodes share the same Cells reference bad things can happen.
        // either never assign to this directly to replace it 
        // or always assign using ShallowCopy() a ShapeCells() object
        public ShapeCells Cells { get; set; }
        
        public List<VA.Text.TabStop> TabStops { get; set; }
        public string CharFontName { get; set; }
        
        protected BaseShape()
        {
            this.Cells = new ShapeCells();
        }
    }
}