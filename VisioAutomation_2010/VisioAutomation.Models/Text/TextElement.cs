using System.Collections.Generic;
using System.Linq;
using VisioAutomation.ShapeSheet.Update;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Text
{
    public class TextElement : Node
    {
        public CharacterFormatting CharacterFormatting { get; set; }
        public ParagraphFormatting ParagraphFormatting { get; set; }

        public TextElement() :
            base(NodeType.Element)
        {
            this.CharacterFormatting = new CharacterFormatting();
            this.ParagraphFormatting = new ParagraphFormatting();
        }

        public TextElement(string text) :
            base(NodeType.Element)
        {
            this.CharacterFormatting = new CharacterFormatting();
            this.ParagraphFormatting = new ParagraphFormatting();
            this.AddText(text);
        }

        public Literal AddText(string text)
        {
            var text_node = new Literal(text);
            this.Add(text_node);
            return text_node;
        }

        public Field AddField(Field field)
        {
            this.Add(field);
            return field;
        }

        public TextElement AddElement()
        {
            var el = new TextElement();
            this.Add(el);
            return el;
        }

        public TextElement AddElement(string text)
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

            foreach (var walkevent in this.Walk())
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
                            var field_region = new TextRegion(start_pos,f);
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
                    string msg = "internal error";
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
                throw new System.ArgumentNullException(nameof(shape));
            }

            // First just set all the text
            string full_doc_inner_text = this.GetInnerText();
            shape.Text = full_doc_inner_text;

            // Find all the regions needing formatting
            var markupinfo = this.GetMarkupInfo();
            var regions_to_format = markupinfo.FormatRegions.Where(region => region.Length >= 1).ToList();

            
            var default_chars_bias = IVisio.VisCharsBias.visBiasLeft;


            var update = new Update();

            foreach (var region in regions_to_format)
            {

                // Apply character formatting
                var charcells = region.Element.CharacterFormatting;
                if (charcells != null)
                {
                    var chars = shape.Characters;
                    chars.Begin = region.Start;
                    chars.End = region.End;
                    chars.CharProps[ShapeSheet.SRCConstants.CharColor.Cell] = 0;
                    short rownum = chars.CharPropsRow[(short) default_chars_bias];

                    if (rownum < 0)
                    {
                        throw new AutomationException("Could not create Character row");
                    }

                    update.Clear();
                    charcells.ApplyFormulas(update, rownum);
                    update.Execute(shape);
                }

                // Apply paragraph formatting
                var paracells = region.Element.ParagraphFormatting;
                if (paracells != null)
                {
                    var chars = shape.Characters;
                    chars.Begin = region.Start;
                    chars.End = region.End;
                    chars.ParaProps[ShapeSheet.SRCConstants.Para_Bullet.Cell] = 0;
                    short rownum = chars.ParaPropsRow[(short) default_chars_bias];

                    if (rownum < 0)
                    {
                        throw new AutomationException("Could not create Paragraph row");
                    }

                    update.Clear();
                    paracells.ApplyFormulas(update, rownum);
                    update.Execute(shape);
                }
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
    }
}