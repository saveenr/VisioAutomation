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
            this.Children.Add(text_node);
            return text_node;
        }

        public Field AppendField(VA.Text.Markup.Field field)
        {
            this.Children.Add(field);
            return field;
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

            var markupinfo = this.GetMarkupInfo();

            string full_doc_inner_text = this.GetInnerText();

            shape.Text = full_doc_inner_text;

            // Format the regions
            foreach (var markup_region in markupinfo.FormatRegions.Where(region => region.TextLength >= 1))
            {
                set_text_range_markup(shape, markup_region);
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

        private static void set_text_range_markup(IVisio.Shape shape, TextRegion markup_region)
        {
            if (markup_region.TextLength < 1)
            {
                return;
            }

            set_text_range_char_fmt(markup_region, shape);
            set_text_range_para_fmt(markup_region, shape);
        }

        private static IVisio.Characters SetRangeParagraphProps(IVisio.Shape shape, short cell, int value, int begin,
                                                 int end)
        {
            var chars = shape.Characters;
            chars.Begin = begin;
            chars.End = end;
            chars.ParaProps[cell] = (short)value;
            return chars;
        }

        private static void set_text_range_para_fmt(TextRegion markup_region, IVisio.Shape shape)
        {
            if (markup_region.Element.ParagraphFormat.Indent.HasValue)
            {
                var chars0 = SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentFirst,
                                                                  0, markup_region.TextStartPos, markup_region.TextEndPos);

                var chars1 = SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentLeft,
                                                                  (int)
                                                                  markup_region.Element.ParagraphFormat.Indent.Value, markup_region.TextStartPos, markup_region.TextEndPos);
            }

            if (markup_region.Element.ParagraphFormat.HAlign.HasValue)
            {
                int int_halign = (int)markup_region.Element.ParagraphFormat.HAlign.Value;
                SetRangeParagraphProps(shape,
                                                     (short) IVisio.VisCellIndices.visHorzAlign,
                                                     int_halign, markup_region.TextStartPos, markup_region.TextEndPos);
            }

            // Handle bullets
            if (markup_region.Element.ParagraphFormat.Bullets.HasValue &&
                markup_region.Element.ParagraphFormat.Bullets.Value)
            {
                const int bullet_type = 1;
                const int base_indent_size = 25;
                int indent_first = -base_indent_size;
                int indent_left = base_indent_size;

                var chars0 = SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentFirst,
                                                                  indent_first, markup_region.TextStartPos, markup_region.TextEndPos);
                var chars1 = SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentLeft,
                                                                  indent_left, markup_region.TextStartPos, markup_region.TextEndPos);
                var chars2 = SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visBulletIndex,
                                                                  bullet_type, markup_region.TextStartPos, markup_region.TextEndPos);
            }
        }

        internal enum rangetype
        {
            Paragraph,
            Character
        }

        internal static void SetRangeProps<T>(IVisio.Shape shape, VA.ShapeSheet.CellData<T> f,
                                              IVisio.VisCellIndices cell, int value, int begin, int end,
                                              ref short rownum, ref IVisio.Characters chars, rangetype rt)
        {

            // http://office.microsoft.com/en-us/visio-help/HV080350454.aspx

            if (!f.Formula.HasValue)
            {
                return;
            }

            var default_chars_bias = IVisio.VisCharsBias.visBiasLeft;
            chars = shape.Characters;
            chars.Begin = begin;
            chars.End = end;

            if (rt == rangetype.Character)
            {
                chars.CharProps[(short)cell] = (short)value;
                rownum = chars.CharPropsRow[(short)default_chars_bias];
            }
            else if (rt == rangetype.Paragraph)
            {
                chars.ParaProps[(short)cell] = (short)value;
                rownum = chars.ParaPropsRow[(short)default_chars_bias];
            }
            else
            {
                throw new System.ArgumentOutOfRangeException("rangetype");
            }

            if (rownum < 0)
            {
                throw new VA.AutomationException("Failed to create a new row");
            }
        }

        private static void set_text_range_char_fmt(TextRegion markup_region, IVisio.Shape shape)
        {
            int startpos = markup_region.TextStartPos;
            int endpos = markup_region.TextEndPos;

            var fmt = new VA.Text.CharacterFormatCells();

            if (markup_region.Element.CharacterFormat.FontSize.HasValue)
            {
                fmt.Size = Convert.PointsToInches(markup_region.Element.CharacterFormat.FontSize.Value);
            }

            if (markup_region.Element.CharacterFormat.Color.HasValue)
            {
                fmt.Color = markup_region.Element.CharacterFormat.Color.Value.ToFormula();
            }

            if (markup_region.Element.CharacterFormat.Font!=null)
            {
                fmt.Font = shape.Document.Fonts[markup_region.Element.CharacterFormat.Font].ID;
            }

            if (markup_region.Element.CharacterFormat.CharStyle.HasValue)
            {
                fmt.Style = (int) markup_region.Element.CharacterFormat.CharStyle.Value;
            }

            if (markup_region.Element.CharacterFormat.Transparency.HasValue)
            {
                fmt.Transparency = markup_region.Element.CharacterFormat.Transparency.Value/100.0;
            }

            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }


            short rownum = -1;
            IVisio.Characters chars = null;

            const int temp_color = 0;
            const int temp_size = 10;
            const int temp_font = 0;
            const int temp_style = 0;
            const int temp_trans = 0;

            SetRangeProps(shape, fmt.Color, IVisio.VisCellIndices.visCharacterColor, temp_color,
                                             startpos, endpos, ref rownum, ref chars, rangetype.Character);
            SetRangeProps(shape, fmt.Size, IVisio.VisCellIndices.visCharacterSize, temp_size,
                                             startpos, endpos, ref rownum, ref chars, rangetype.Character);
            SetRangeProps(shape, fmt.Font, IVisio.VisCellIndices.visCharacterFont, temp_font,
                                             startpos, endpos, ref rownum, ref chars, rangetype.Character);
            SetRangeProps(shape, fmt.Style, IVisio.VisCellIndices.visCharacterStyle, temp_style,
                                             startpos, endpos, ref rownum, ref chars, rangetype.Character);
            SetRangeProps(shape, fmt.Transparency, IVisio.VisCellIndices.visCharacterColorTrans,
                                             temp_trans, startpos, endpos, ref rownum, ref chars,
                                             rangetype.Character);
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