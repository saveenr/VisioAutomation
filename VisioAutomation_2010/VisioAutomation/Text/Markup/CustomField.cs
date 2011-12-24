using IVisio = Microsoft.Office.Interop.Visio;
using System;
using VA=VisioAutomation;
using System.Linq;

namespace VisioAutomation.Text.Markup
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
