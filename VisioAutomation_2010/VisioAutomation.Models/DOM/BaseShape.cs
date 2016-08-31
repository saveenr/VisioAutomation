using System.Collections.Generic;
using VisioAutomation.Text;
using VACUSTPROP = VisioAutomation.Shapes.CustomProperties;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Dom
{
    public class BaseShape : Node
    {
        public IVisio.Shape VisioShape { get; set; }
        public short VisioShapeID { get; internal set; }

        public VisioAutomation.Models.Text.TextElement Text { get; set; }
        public VisioAutomation.Shapes.CustomProperties.CustomPropertyDictionary CustomProperties { get; set; }
        public List<Hyperlink> Hyperlinks { get; set; }

        // Be aware that if multiple nodes share the same Cells reference bad things can happen.
        // either never assign to this directly to replace it 
        // or always assign using ShallowCopy() a ShapeCells() object
        public ShapeCells Cells { get; set; }
        
        public List<TabStop> TabStops { get; set; }
        public string CharFontName { get; set; }
        
        protected BaseShape()
        {
            this.Cells = new ShapeCells();
        }
    }
}