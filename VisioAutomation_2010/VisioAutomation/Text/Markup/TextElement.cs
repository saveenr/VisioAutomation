using System;
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

                var charcells = markup_region.Element.CharacterFormat.ToCells();
                FormatTextRegion(shape, charcells, markup_region.Start, markup_region.End); 

                var paracells = markup_region.Element.ParagraphFormat.ToCells();
                FormatTextRegion(shape, paracells, markup_region.Start, markup_region.End);
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

        public static int FormatTextRegion(IVisio.Shape shape, VA.ShapeSheet.CellGroups.CellGroupMultiRow fmtcells, int start, int end)
        {
            // overall strategy:
            // we need a row (either an existing one or a new one to be created)
            // once we have a row, we can set the cells like normal for that row

            // Ensure that this method only works on character and paragraph cells
            if (!(fmtcells is CharacterFormatCells || fmtcells is ParagraphFormatCells))
            {
                string msg = string.Format("Only accepts {0} or {1}", typeof (CharacterFormatCells).Name,
                                           typeof (ParagraphFormatCells).Name);
                throw new VA.AutomationException(msg);
            }

            // ensure we have a valid shape object
            if (shape == null)
            {
                throw new System.ArgumentNullException("shape");
            }

            // Initialize the properties with temp values
            short rownum = -1;
            var default_chars_bias = IVisio.VisCharsBias.visBiasLeft;

            // Try to create either a character or paragraph row depending on the cells that were passed into this method
            IVisio.Characters chars = shape.Characters;
            chars.Begin = start;
            chars.End = end;

            if (fmtcells is CharacterFormatCells)
            {
                // the choice of Color arbitrary 
                chars.CharProps[SRCCON.Char_Color.Cell] = (short)0;
                rownum = chars.CharPropsRow[(short)default_chars_bias];                
            }
            else if (fmtcells is ParagraphFormatCells)
            {
                // the choice of Bullet is arbitrary
                chars.ParaProps[SRCCON.Para_Bullet.Cell] = (short)0;
                rownum = chars.ParaPropsRow[(short)default_chars_bias];                
            }
            else
            {
                throw new ArgumentOutOfRangeException("fmtcells");
            }

            // If a negative rownum was return the reason is that the desired new region spanned multiple existing regions
            if (rownum == -1)
            {
                throw new VA.AutomationException("Cannot apply formatting across multiple regions");
            }

            // Now that we have a row identified apply the cells
            var update = new VA.ShapeSheet.Update.SRCUpdate();
            fmtcells.Apply(update, rownum);
            update.Execute(shape);

            // return the rownumber in case the caller wants to do something with the row that the formatting is on
            return rownum;
        }
    }
}