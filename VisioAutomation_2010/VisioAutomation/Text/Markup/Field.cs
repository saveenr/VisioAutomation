using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text.Markup
{
    public class Field : Node
    {
        private const string placeholder_string = "[FIELD]";

        public IVisio.VisFieldCategories Category { get; set; }
        public IVisio.VisFieldCodes Code { get; set; }
        public IVisio.VisFieldFormats Format { get; set; }

        public Field(IVisio.VisFieldCategories category, IVisio.VisFieldCodes code, IVisio.VisFieldFormats format) :
            base(NodeType.Field)
        {
            this.Category = category;
            this.Code = code;
            this.Format = format;
        }

        public string PlaceholderText
        {
            get
            {
                return placeholder_string;
            }
        }
    }
}
