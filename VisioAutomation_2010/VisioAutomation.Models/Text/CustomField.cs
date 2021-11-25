
namespace VisioAutomation.Models.Text
{
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
