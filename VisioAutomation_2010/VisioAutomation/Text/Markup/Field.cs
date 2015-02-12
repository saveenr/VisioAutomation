using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Text.Markup
{
    public class Field : FieldBase
    {
        public IVisio.VisFieldCategories Category { get; set; }
        public IVisio.VisFieldCodes Code { get; set; }

        public Field(IVisio.VisFieldCategories category, IVisio.VisFieldCodes code, IVisio.VisFieldFormats format) :
            base(NodeType.Field)
        {
            this.Category = category;
            this.Code = code;
            this.Format = format;
        }
    }
}
