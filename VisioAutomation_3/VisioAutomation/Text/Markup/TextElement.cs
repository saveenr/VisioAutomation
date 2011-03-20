using System.Collections.Generic;
using System.Linq;
using System.Xml;
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

        public Field AppendField(IVisio.VisFieldCategories category, IVisio.VisFieldCodes code,
                                 IVisio.VisFieldFormats format)
        {
            var f = new Field(category, code, format);
            this.Children.Add(f);
            return f;
        }

        public TextElement AppendNewElement()
        {
            var el = new TextElement();
            this.Children.Add(el);
            return el;
        }

        public TextElement AppendNewElement(string text)
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

        public static TextElement FromXml(string input_xml, bool preserve_whitespace)
        {
            

            System.Xml.Linq.LoadOptions lo = System.Xml.Linq.LoadOptions.None;
            if (preserve_whitespace)
            {
                lo = lo & System.Xml.Linq.LoadOptions.PreserveWhitespace;
            }

            var xml_doc = System.Xml.Linq.XDocument.Parse(input_xml, lo);

            var text_el =
                (TextElement)create_va_text_node_from_xml_node(xml_doc.Root, preserve_whitespace);

            return text_el;
        }

        private static IEnumerable<System.Xml.Linq.XNode> get_child_nodes(System.Xml.Linq.XNode node)
        {
            if (node is System.Xml.Linq.XElement)
            {
                var node_el = (System.Xml.Linq.XElement) node;
                foreach (var i in node_el.Nodes())
                {
                    yield return i;
                }
            }
            else
            {
                yield break;
            }
        }

        private static IEnumerable<VA.Internal.WalkEvent<System.Xml.Linq.XNode>> walk_xml_node(System.Xml.Linq.XNode node)
        {
            return VA.Internal.TreeTraversal.Walk(node, n => get_child_nodes(n), n => true);
        }

        private static Node create_va_text_node_from_xml_node(System.Xml.Linq.XNode node, bool preserve_whitespace)
        {
            var root_el = new TextElement();

            var stack = new Stack<TextElement>();
            stack.Push(root_el);

            foreach (var walkevent in walk_xml_node(node))
            {
                if (walkevent.HasEnteredNode)
                {
                    fromxml_enter_node(walkevent.Node, stack, preserve_whitespace);
                }
                else if (walkevent.HasExitedNode)
                {
                    fromxml_exit_node(walkevent.Node, stack);
                }
            }

            return root_el;
        }

        private static void fromxml_enter_node(System.Xml.Linq.XNode node, Stack<TextElement> stack,
                                               bool preserve_whitespace)
        {

            if (node is System.Xml.Linq.XElement)
            {
                var node_el = (System.Xml.Linq.XElement) node;

                if (node_el.Name == "text")
                {
                    var current_el = new TextElement();
                    current_el.TextFormat.LoadAttributesFromXml(node_el);
                    stack.Push(current_el);
                }
                else if ((node_el.Name == "br") || (node_el.Name == "newline"))
                {
                    var parent = stack.Peek();
                    parent.AppendText("\n");
                }
                else if (node_el.Name == "tab")
                {
                    var parent = stack.Peek();
                    parent.AppendText("\t");
                }
                else if (node_el.Name == "space")
                {
                    var parent = stack.Peek();
                    parent.AppendText(" ");
                }
                else
                {
                    string msg = string.Format("unsupported element {0}", node_el.Name);
                    throw new System.ArgumentException("node", msg);
                }
            }
            else if (node is System.Xml.Linq.XText)
            {
                // These nodes contribute text, so update the current region
                var parent = stack.Peek();

                var node_text = (System.Xml.Linq.XText) node;
                string t = node_text.Value;
                if (!preserve_whitespace)
                {
                    t = t.Trim();
                }

                parent.AppendText(t);
            }
            else if (node is System.Xml.Linq.XComment)
            {
                //do nothing
            }
            else if (node is System.Xml.Linq.XProcessingInstruction)
            {
                //do nothing
            }
            else
            {
                string msg = string.Format("Unhandled node type {0}", node.GetType());
                throw new System.ArgumentOutOfRangeException("Node", msg);
            }
        }


        private static void fromxml_exit_node(System.Xml.Linq.XNode node, Stack<TextElement> stack)
        {
            if (node is System.Xml.Linq.XElement)
            {
                var node_el = (System.Xml.Linq.XElement)node;

                if (node_el.Name == "text")
                {
                    var current_el = stack.Pop();
                    var parent_el = stack.Peek();
                    parent_el.Children.Add(current_el);
                }
                
            }
        }

        public MarkupInfo GetMarkupInfo()
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

        public void SetShapeText(IVisio.Shape shape)
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
            foreach (var field_region in markupinfo.FieldRegions.Where(region => region.TextLength >= 1))
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

        private static void set_text_range_para_fmt(TextRegion markup_region, IVisio.Shape shape)
        {
            if (markup_region.Element.TextFormat.Indent.HasValue)
            {
                var chars0 = VA.Text.TextHelper.SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentFirst,
                                                                  0, markup_region.TextStartPos, markup_region.TextEndPos);

                var chars1 = VA.Text.TextHelper.SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentLeft,
                                                                  (int)
                                                                  markup_region.Element.TextFormat.Indent.Value, markup_region.TextStartPos, markup_region.TextEndPos);
            }

            if (markup_region.Element.TextFormat.HAlign.HasValue)
            {
                int int_halign = (int) markup_region.Element.TextFormat.HAlign.Value;
                VA.Text.TextHelper.SetRangeParagraphProps(shape,
                                                     (short) IVisio.VisCellIndices.visHorzAlign,
                                                     int_halign, markup_region.TextStartPos, markup_region.TextEndPos);
            }

            // Handle bullets
            if (markup_region.Element.TextFormat.Bullets.HasValue &&
                markup_region.Element.TextFormat.Bullets.Value)
            {
                const int bullet_type = 1;
                const int base_indent_size = 25;
                int indent_first = -base_indent_size;
                int indent_left = base_indent_size;

                var chars0 = VA.Text.TextHelper.SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentFirst,
                                                                  indent_first, markup_region.TextStartPos, markup_region.TextEndPos);
                var chars1 = VA.Text.TextHelper.SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visIndentLeft,
                                                                  indent_left, markup_region.TextStartPos, markup_region.TextEndPos);
                var chars2 = VA.Text.TextHelper.SetRangeParagraphProps(shape,
                                                                  (short) IVisio.VisCellIndices.visBulletIndex,
                                                                  bullet_type, markup_region.TextStartPos, markup_region.TextEndPos);
            }
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

            VA.Text.TextHelper.SetFormat(shape, fmt, startpos, endpos);
        }
    }
}