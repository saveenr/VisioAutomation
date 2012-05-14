using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class BaseShape : Node
    {
        public IVisio.Shape VisioShape { get; set; }
        public short VisioShapeID { get; internal set; }

        public VA.Text.Markup.TextElement Text { get; set; }
        public Dictionary<string, VA.CustomProperties.CustomPropertyCells> CustomProperties { get; set; }
        public List<Hyperlink> Hyperlinks { get; set; }
        public ShapeCells Cells { get; set; }
        public List<VA.Text.TabStop> TabStops { get; set; }
        public string CharFontName { get; set; }
        
        protected BaseShape()
        {
            this.Cells = new ShapeCells();
        }
    }
}