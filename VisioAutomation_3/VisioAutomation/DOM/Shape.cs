using System.Collections.Generic;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.DOM
{
    public class Shape : Node
    {
        public string Text { get; set; }
        public VA.Text.Markup.TextElement TextElement { get; set; }

        public Dictionary<string, VA.CustomProperties.CustomPropertyCells> CustomProperties { get; set; }
        public List<Hyperlink> Hyperlinks { get; set; }
        public ShapeCells ShapeCells { get; set; }
        public List<VA.Text.TabStop> TabStops { get; set; }
        public IVisio.Shape VisioShape;
        public short VisioShapeID { get; internal set; }
        public string CharFontName { get; set; }
        
        protected Shape()
        {
            this.ShapeCells = new ShapeCells();
        }

        public VA.CustomProperties.CustomPropertyCells SetCustomProperty(string name, string value)
        {
            var cp = new VA.CustomProperties.CustomPropertyCells();
            cp.Value = value;

            if (this.CustomProperties == null)
            {
                this.CustomProperties = new Dictionary<string, VA.CustomProperties.CustomPropertyCells>();
            }

            this.CustomProperties[name] = cp;
            return cp;
        }
    }
}