using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

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
                        var region = new TextRegion(element, start_pos);
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
                            nparent.TextLength += text_node.Text.Length;

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
                            field_region.TextStartPos = start_pos;
                            field_region.TextLength = f.PlaceholderText.Length;

                            markupinfo.FieldRegions.Add(field_region);

                            // Add text length to parent
                            var nparent = region_stack.Peek();
                            nparent.TextLength += f.PlaceholderText.Length;

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

            // First just set all the text
            string full_doc_inner_text = this.GetInnerText();
            shape.Text = full_doc_inner_text;

            // Find all the regions needing formatting
            var markupinfo = this.GetMarkupInfo();
            var regions_to_format = markupinfo.FormatRegions.Where(region => region.TextLength >= 1);
            foreach (var markup_region in regions_to_format)
            {
                set_text_range_char_fmt(markup_region, shape);
                set_text_range_para_fmt(markup_region, shape);
            }

            // Insert the fields
            // note: Fields are added in reverse because it is simpler to keep track of the insertion positions
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


        private static void set_text_range_para_fmt(TextRegion region, IVisio.Shape shape)
        {
            if (region.Element.ParagraphFormat.IndentFirst.HasValue)
            {
                int indent_first_points = (int)VA.Convert.InchestoPoints(region.Element.ParagraphFormat.IndentFirst.Value);
                var chars0 = SetRangeParagraphProps(shape, VA.ShapeSheet.SRCConstants.Para_IndFirst, indent_first_points, region);
            }

            if (region.Element.ParagraphFormat.IndentLeft.HasValue)
            {
                int indent_left_points = (int) VA.Convert.InchestoPoints(region.Element.ParagraphFormat.IndentLeft.Value);
                var chars1 = SetRangeParagraphProps(shape, VA.ShapeSheet.SRCConstants.Para_IndLeft, indent_left_points, region);
            }

            if (region.Element.ParagraphFormat.HAlign.HasValue)
            {
                int int_halign = (int)region.Element.ParagraphFormat.HAlign.Value;
                SetRangeParagraphProps(shape, VA.ShapeSheet.SRCConstants.Para_HorzAlign, int_halign, region);
            }

            // Handle bullets
            if (region.Element.ParagraphFormat.Bullets.HasValue &&
                region.Element.ParagraphFormat.Bullets.Value)
            {
                const int bullet_type = 1;
                const int base_indent_size = 25;
                int indent_first = -base_indent_size;
                int indent_left = base_indent_size;

                var chars0 = SetRangeParagraphProps(shape, VA.ShapeSheet.SRCConstants.Para_IndFirst, indent_first, region);
                var chars1 = SetRangeParagraphProps(shape, VA.ShapeSheet.SRCConstants.Para_IndLeft, indent_left, region);
                var chars2 = SetRangeParagraphProps(shape, VA.ShapeSheet.SRCConstants.Para_Bullet, bullet_type, region);
            }
        }

        internal enum rangetype
        {
            Paragraph,
            Character
        }

        private static IVisio.Characters SetRangeParagraphProps(IVisio.Shape shape, VA.ShapeSheet.SRC src, int value, VA.Text.Markup.TextRegion region)
        {
            var chars = shape.Characters;
            chars.Begin = region.TextStartPos;
            chars.End = region.TextEndPos;
            chars.ParaProps[src.Cell] = (short)value;
            return chars;
        }

        private static void SetRangeProps<T>(IVisio.Shape shape, VA.ShapeSheet.CellData<T> f,
                                      VA.ShapeSheet.SRC src, int value, Markup.TextRegion region,
                                      ref short rownum, ref IVisio.Characters chars)
        {
            // http://office.microsoft.com/en-us/visio-help/HV080350454.aspx

            if (!f.Formula.HasValue)
            {
                return;
            }

            var default_chars_bias = IVisio.VisCharsBias.visBiasLeft;
            chars = shape.Characters;
            chars.Begin = region.TextStartPos;
            chars.End = region.TextEndPos;

            if (src.Section == (short)IVisio.VisSectionIndices.visSectionCharacter)
            {
                chars.CharProps[src.Cell] = (short)value;
                rownum = chars.CharPropsRow[(short)default_chars_bias];
            }
            else if (src.Section == (short)IVisio.VisSectionIndices.visSectionParagraph)
            {
                chars.ParaProps[src.Cell] = (short)value;
                rownum = chars.ParaPropsRow[(short)default_chars_bias];
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("rangetype");
            }

            if (rownum < 0)
            {
                throw new VA.AutomationException("Failed to create a new row. Because range spanned multiple existing rows");
            }
        }

        private static void set_text_range_char_fmt(TextRegion region, IVisio.Shape shape)
        {
            var fmt = new VA.Text.CharacterFormatCells();

            if (region.Element.CharacterFormat.Size.HasValue)
            {
                fmt.Size = Convert.PointsToInches(region.Element.CharacterFormat.Size.Value);
            }

            if (region.Element.CharacterFormat.Color.HasValue)
            {
                fmt.Color = region.Element.CharacterFormat.Color.Value.ToFormula();
            }

            if (region.Element.CharacterFormat.Font!=null)
            {
                var doc = shape.Document;
                var fonts = doc.Fonts;
                var font = fonts[region.Element.CharacterFormat.Font];
                fmt.Font = font.ID;
            }

            if (region.Element.CharacterFormat.Style.HasValue)
            {
                fmt.Style = (int) region.Element.CharacterFormat.Style.Value;
            }

            if (region.Element.CharacterFormat.Transparency.HasValue)
            {
                fmt.Transparency = region.Element.CharacterFormat.Transparency.Value/100.0;
            }

            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            // Initialize the properties with temp values
            short rownum = -1;
            const int temp_color = 0;
            const int temp_size = 10;
            const int temp_font = 0;
            const int temp_style = 0;
            const int temp_trans = 0;

            IVisio.Characters chars = null;
            SetRangeProps(shape, fmt.Color, VA.ShapeSheet.SRCConstants.Char_Color, temp_color, region, ref rownum, ref chars);
            SetRangeProps(shape, fmt.Size, VA.ShapeSheet.SRCConstants.Char_Size, temp_size, region, ref rownum, ref chars);
            SetRangeProps(shape, fmt.Font, VA.ShapeSheet.SRCConstants.Char_Font, temp_font, region, ref rownum, ref chars);
            SetRangeProps(shape, fmt.Style, VA.ShapeSheet.SRCConstants.Char_Style, temp_style, region, ref rownum, ref chars);
            SetRangeProps(shape, fmt.Transparency, VA.ShapeSheet.SRCConstants.Char_ColorTrans, temp_trans, region, ref rownum, ref chars);

            // if any text region was created then set the formula values
            if (chars != null)
            {
                if (rownum < 0)
                {
                    throw new AutomationException("Internal Error");
                }

                var update = new VA.ShapeSheet.Update.SRCUpdate();
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Color.ForRow(rownum), fmt.Color.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Size.ForRow(rownum), fmt.Size.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Font.ForRow(rownum), fmt.Font.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_Style.ForRow(rownum), fmt.Style.Formula);
                update.SetFormulaIgnoreNull(VA.ShapeSheet.SRCConstants.Char_ColorTrans.ForRow(rownum), fmt.Transparency.Formula);
                update.Execute(shape);
            }
        }
    }
}