using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Text.Markup
{
    public class TextElement : Node
    {
        public TextElement() :
            base(NodeType.Element)
        {
            this.TextFormat = new TextFormat();
        }

        public TextElement(string text) :
            base(NodeType.Element)
        {
            this.TextFormat = new TextFormat();
            this.AppendText(text);
        }

        public Literal AppendText(string text)
        {
            var text_node = new Literal(text);
            this.Children.Add(text_node);
            return text_node;
        }

        public Field AppendField(VA.Text.Markup.Field f)
        {
            this.Children.Add(f);
            return f;
        }

        public TextElement AppendElement()
        {
            var el = new TextElement();
            this.Children.Add(el);
            return el;
        }

        public TextElement AppendElement(string text)
        {
            var el = new TextElement(text);
            this.Children.Add(el);
            return el;
        }

        public IEnumerable<TextElement> Elements
        {
            get { return this.Children.Items.Where(n => n.NodeType == NodeType.Element).Cast<TextElement>(); }
        }

        public TextFormat TextFormat { get; set; }


        internal MarkupInfo GetMarkupInfo()
        {
            var markupinfo = new MarkupInfo();

            int start_pos = 0;
            var region_stack = new Stack<TextRegion>();

            foreach (var walkevent in Walk())
            {
                if (walkevent.HasEnteredNode)
                {
                    if (walkevent.Node is TextElement)
                    {
                        var element = (TextElement) walkevent.Node;
                        var region = new TextRegion(element, start_pos);
                        region_stack.Push(region);
                        markupinfo.FormatRegions.Add(region);
                    }
                    if (walkevent.Node is Literal)
                    {
                        var text_node = (Literal) walkevent.Node;

                        if (!string.IsNullOrEmpty(text_node.Text))
                        {
                            // Add text length to parent
                            var nparent = region_stack.Peek();
                            nparent.TextLength += text_node.Text.Length;

                            // update the start position with the length
                            start_pos += text_node.Text.Length;
                        }
                    }
                    else if (walkevent.Node is Field)
                    {
                        var field = (Field) walkevent.Node;
                        if (field.PlaceholderText == null)
                        {
                            string msg = "Placeholder text cannot be null";
                            throw new VA.AutomationException(msg);
                        }

                        var field_region = new TextRegion();
                        field_region.Field = field;
                        field_region.TextStartPos = start_pos;
                        field_region.TextLength = field.PlaceholderText.Length;

                        markupinfo.FieldRegions.Add(field_region);

                        // Add text length to parent
                        var nparent = region_stack.Peek();
                        nparent.TextLength += field.PlaceholderText.Length;

                        // update the start position with the length
                        start_pos += field.PlaceholderText.Length;
                    }
                    else
                    {
                        // do nothing
                    }
                }
                else if (walkevent.HasExitedNode)
                {
                    if (walkevent.Node is TextElement)
                    {
                        var this_region = region_stack.Pop();

                        if (region_stack.Count > 0)
                        {
                            var parent_el = region_stack.Peek();
                            parent_el.TextLength += this_region.TextLength;
                        }
                    }
                }
                else
                {
                    // Unhandled Operation
                    string msg = string.Format("internal error");
                    throw new System.InvalidOperationException(msg);
                }
            }

            if (region_stack.Count > 0)
            {
                throw new System.InvalidOperationException("Regions stack not empty");
            }

            return markupinfo;
        }

        public void SetText(IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            var markupinfo = this.GetMarkupInfo();

            string full_doc_inner_text = this.GetInnerText();

            shape.Text = full_doc_inner_text;

            // Format the regions
            foreach (var markup_region in markupinfo.FormatRegions.Where(region => region.TextLength >= 1))
            {
                set_text_range_char_fmt(markup_region, shape);
                set_text_range_para_fmt(markup_region, shape);
            }

            // Insert the fields
            // NOTE - to keep it simpler the fields are added backwards - last to first
            foreach (var field_region in markupinfo.FieldRegions.Where(region => region.TextLength >= 1).Reverse())
            {
                var chars = shape.Characters;
                chars.Begin = field_region.TextStartPos;
                chars.End = field_region.TextEndPos;
                chars.AddField((short) field_region.Field.Category, (short) field_region.Field.Code,
                               (short) field_region.Field.Format);
                var fr = field_region;
            }
        }


        private static void set_text_range_para_fmt(TextRegion markup_region, IVisio.Shape shape)
        {
            int startpos = markup_region.TextStartPos;
            int endpos = markup_region.TextEndPos;

            var fmt = new VA.Text.ParagraphFormatCells();
            var tf = markup_region.Element.TextFormat;
            
            if (tf.Indent.HasValue)
            {
                fmt.IndentFirst = VA.Convert.PointsToInches(tf.Indent.Value);
                fmt.IndentLeft = VA.Convert.PointsToInches(tf.Indent.Value);
            }

            if (tf.HAlign.HasValue)
            {
                fmt.HorizontalAlign = (int) tf.HAlign.Value;
            }

            // Handle bullets
            if (tf.Bullets.HasValue &&
                tf.Bullets.Value)
            {
                const int base_indent_size = 25;
                fmt.IndentFirst = "-25pt";
                fmt.IndentLeft = "25pt";
                fmt.BulletIndex = "1";
            }
            
            VA.Text.TextFormat.FormatRange(shape, fmt, startpos, endpos);
        }

        private static void set_text_range_char_fmt(TextRegion markup_region, IVisio.Shape shape)
        {
            int startpos = markup_region.TextStartPos;
            int endpos = markup_region.TextEndPos;

            var fmt = new VA.Text.CharacterFormatCells();

            if (markup_region.Element.TextFormat.FontSize.HasValue)
            {
                fmt.Size = Convert.PointsToInches(markup_region.Element.TextFormat.FontSize.Value);
            }

            if (markup_region.Element.TextFormat.Color.HasValue)
            {
                fmt.Color = markup_region.Element.TextFormat.Color.Value.ToFormula();
            }

            if (markup_region.Element.TextFormat.Font!=null)
            {
                fmt.Font = shape.Document.Fonts[markup_region.Element.TextFormat.Font].ID;
            }

            if (markup_region.Element.TextFormat.CharStyle.HasValue)
            {
                fmt.Style = (int) markup_region.Element.TextFormat.CharStyle.Value;
            }

            if (markup_region.Element.TextFormat.Transparency.HasValue)
            {
                fmt.Transparency = markup_region.Element.TextFormat.Transparency.Value/100.0;
            }

            VA.Text.TextFormat.FormatRange(shape, fmt, startpos, endpos);
        }
    }
}