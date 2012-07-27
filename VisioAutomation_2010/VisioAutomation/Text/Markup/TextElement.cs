using System.Collections.Generic;
using System.Linq;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Text.Markup
{
    public class TextElement : Node
    {
        public CharacterFormat CharacterFormat { get; set; }
        public ParagraphFormat ParagraphFormat { get; set; }

        public TextElement() :
            base(NodeType.Element)
        {
            this.CharacterFormat = new CharacterFormat();
            this.ParagraphFormat  = new ParagraphFormat();
        }

        public TextElement(string text) :
            base(NodeType.Element)
        {
            this.CharacterFormat = new CharacterFormat();
            this.ParagraphFormat = new ParagraphFormat();
            this.AppendText(text);
        }

        public Literal AppendText(string text)
        {
            var text_node = new Literal(text);
            this.Add(text_node);
            return text_node;
        }

        public Field AppendField(VA.Text.Markup.Field field)
        {
            this.Add(field);
            return field;
        }

        public TextElement AppendElement()
        {
            var el = new TextElement();
            this.Add(el);
            return el;
        }

        public TextElement AppendElement(string text)
        {
            var el = new TextElement(text);
            this.Add(el);
            return el;
        }

        public IEnumerable<TextElement> Elements
        {
            get { return this.Children.Where(n => n.NodeType == NodeType.Element).Cast<TextElement>(); }
        }
        
        internal MarkupRegions GetMarkupInfo()
        {
            var markupinfo = new MarkupRegions();

            int start_pos = 0;
            var region_stack = new Stack<TextRegion>();

            foreach (var walkevent in Walk())
            {
                if (walkevent.HasEnteredNode)
                {
                    if (walkevent.Node is TextElement)
                    {
                        var element = (TextElement) walkevent.Node;
                        var region = new TextRegion(start_pos, element);
                        region_stack.Push(region);
                        markupinfo.FormatRegions.Add(region);
                    }
                    else if (walkevent.Node is Literal)
                    {
                        var text_node = (Literal) walkevent.Node;

                        if (!string.IsNullOrEmpty(text_node.Text))
                        {
                            // Add text length to parent
                            var nparent = region_stack.Peek();
                            nparent.Length += text_node.Text.Length;

                            // update the start position with the length
                            start_pos += text_node.Text.Length;
                        }
                    }
                    else if (walkevent.Node is Field)
                    {
                        var f = (Field) walkevent.Node;
                        if (!string.IsNullOrEmpty(f.PlaceholderText))
                        {
                            var field_region = new TextRegion();
                            field_region.Field = f;
                            field_region.Start = start_pos;
                            field_region.Length = f.PlaceholderText.Length;

                            markupinfo.FieldRegions.Add(field_region);

                            // Add text length to parent
                            var nparent = region_stack.Peek();
                            nparent.Length += f.PlaceholderText.Length;

                            // update the start position with the length
                            start_pos += f.PlaceholderText.Length;
                        }
                    }
                    else
                    {
                        string msg = "Unhandled node";
                        throw new AutomationException(msg);
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
                            parent_el.Length += this_region.Length;
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

            // First just set all the text
            string full_doc_inner_text = this.GetInnerText();
            shape.Text = full_doc_inner_text;

            // Find all the regions needing formatting
            var markupinfo = this.GetMarkupInfo();
            var regions_to_format = markupinfo.FormatRegions.Where(region => region.Length >= 1);
            foreach (var markup_region in regions_to_format)
            {

                var charfmt = markup_region.Element.CharacterFormat;
                var charcells = charfmt.ToCells();
                FormatTextRegion(charcells, markup_region, shape); 
                set_text_range_para_fmt(markup_region, shape);
            }

            // Insert the fields
            // note: Fields are added in reverse because it is simpler to keep track of the insertion positions
            foreach (var field_region in markupinfo.FieldRegions.Where(region => region.Length >= 1).Reverse())
            {
                var chars = shape.Characters;
                chars.Begin = field_region.Start;
                chars.End = field_region.End;
                chars.AddField((short) field_region.Field.Category, (short) field_region.Field.Code,
                               (short) field_region.Field.Format);
                var fr = field_region;
            }
        }


        private static void set_text_range_para_fmt(TextRegion region, IVisio.Shape shape)
        {
            short rownum = -1;
            IVisio.Characters chars=null;

            var parafmt = region.Element.ParagraphFormat;

            if (parafmt.IndentFirstInPoints.HasValue)
            {
                int indent_first_points = (int)VA.Convert.InchestoPoints(parafmt.IndentFirstInPoints.Value);
                var chars0 = SetRangeParagraphProps(shape, parafmt.IndentFirstInPoints.HasValue, SRCCON.Para_IndFirst, indent_first_points, region, ref rownum, ref chars);
            }

            if (parafmt.IndentLeftInPoints.HasValue)
            {
                int indent_left_points = (int) VA.Convert.InchestoPoints(parafmt.IndentLeftInPoints.Value);
                var chars1 = SetRangeParagraphProps(shape, parafmt.IndentLeftInPoints.HasValue, SRCCON.Para_IndLeft, indent_left_points, region, ref rownum, ref chars);
            }

            if (parafmt.HAlign.HasValue)
            {
                int int_halign = (int)parafmt.HAlign.Value;
                SetRangeParagraphProps(shape, parafmt.HAlign.HasValue, SRCCON.Para_HorzAlign, int_halign, region, ref rownum, ref chars);
            }

            // Handle bullets
            if (parafmt.Bullets.HasValue &&
                parafmt.Bullets.Value)
            {
                const int bullet_type = 1;
                const int base_indent_size = 25;
                int indent_first = -base_indent_size;
                int indent_left = base_indent_size;

                var chars0 = SetRangeParagraphProps(shape, parafmt.Bullets.HasValue, SRCCON.Para_IndFirst, indent_first, region, ref rownum, ref chars);
                var chars1 = SetRangeParagraphProps(shape, parafmt.Bullets.HasValue, SRCCON.Para_IndLeft, indent_left, region, ref rownum, ref chars);
                var chars2 = SetRangeParagraphProps(shape, parafmt.Bullets.HasValue, SRCCON.Para_Bullet, bullet_type, region, ref rownum, ref chars);
            }
        }

        private static IVisio.Characters SetRangeParagraphProps(IVisio.Shape shape, bool perform, VA.ShapeSheet.SRC src, int value, VA.Text.Markup.TextRegion region, ref short rownum, ref IVisio.Characters chars2)
        {
            if (!perform)
            {
                return null;
            }

            var chars = shape.Characters;
            chars.Begin = region.Start;
            chars.End = region.End;
            chars.ParaProps[src.Cell] = (short)value;
            return chars;
        }

        private static void FormatTextRegion(CharacterFormatCells charcells, TextRegion region, IVisio.Shape shape)
        {
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            // Initialize the properties with temp values
            short rownum = -1;
            var default_chars_bias = IVisio.VisCharsBias.visBiasLeft;
            IVisio.Characters chars = null;

            chars = shape.Characters;
            chars.Begin = region.Start;
            chars.End = region.End;

            chars.CharProps[SRCCON.Char_Color.Cell] = (short)0;
            rownum = chars.CharPropsRow[(short)default_chars_bias];

            if (rownum == -1)
            {
                throw new VA.AutomationException("Internal Error");
            }

            if (chars == null)
            {
                throw new VA.AutomationException("Internal Error2");

            }

            var update = new VA.ShapeSheet.Update.SRCUpdate();
            charcells.Apply(update, rownum);
            update.Execute(shape);
        }
    }
}