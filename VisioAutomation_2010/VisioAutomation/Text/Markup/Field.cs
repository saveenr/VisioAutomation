using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text.Markup
{
    public class FieldBase : Node
    {
        internal FieldBase(NodeType nt) : base(nt)
        {
        }

        private const string placeholder_string = "[FIELD]";
        public IVisio.VisFieldFormats Format { get; set; }

        public string PlaceholderText
        {
            get
            {
                return placeholder_string;
            }
        }
    }

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

    public class CustomField: FieldBase
    {
        public string Formula { get; set; }

        public CustomField(string formula, IVisio.VisFieldFormats fmt) :
            base(NodeType.Field)
        {
            this.Formula = formula;
            this.Format = fmt;
        }
    }
}
