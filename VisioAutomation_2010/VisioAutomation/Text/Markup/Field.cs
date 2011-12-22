using IVisio = Microsoft.Office.Interop.Visio;
using System;
using VA=VisioAutomation;
using System.Linq;

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

        public static void SetText(IVisio.Shape shape, string fmt, params VA.Text.Markup.FieldBase[] fields)
        {
            if (shape == null)
            {
                throw new ArgumentNullException("shape");
            }

            if (fields == null)
            {
                throw new ArgumentNullException("fields");
            }

            var fmtparse = new VA.Internal.FormatStringParser(fmt);
            var unique_indices = fmtparse.Segments.Select(f => f.Index).Distinct().ToList();
            if (unique_indices.Count > fields.Length)
            {
                throw new ArgumentOutOfRangeException("fmt", "index out of range for number of insertions");
            }

            // Set the text
            shape.Text = fmt;

            // then Insert the fields from last to first (makes it easier to keep track of positions this way)
            for (int i = (fmtparse.Segments.Count - 1); i >= 0; i--)
            {
                var fmt_seg = fmtparse.Segments[i];
                var field_index = fmt_seg.Index;
                var field = fields[field_index];

                var chars = shape.Characters;
                chars.Begin = fmt_seg.Start;
                chars.End = fmt_seg.End;

                if (field is VA.Text.Markup.CustomField)
                {
                    var customfield = (VA.Text.Markup.CustomField)field;
                    chars.AddCustomFieldU(customfield.Formula, (short)customfield.Format);
                }
                else if (field is VA.Text.Markup.Field)
                {
                    var field_f = (VA.Text.Markup.Field)field;
                    chars.AddField((short)field_f.Category, (short)field_f.Code, (short)field_f.Format);
                }
                else
                {
                    string msg = String.Format("Unsupported field type {0} for field {1}", field.GetType(), i);
                    throw new AutomationException(msg);
                }
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
