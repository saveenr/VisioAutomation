
using GenTreeOps;
using VisioAutomation.Exceptions;
using VisioAutomation.ShapeSheet;


namespace VisioAutomation.Models.Text;

public class Element : Node
{
    public CharacterFormatting CharacterFormatting { get; set; }
    public ParagraphFormatting ParagraphFormatting { get; set; }

    public Element() :
        base(NodeType.Element)
    {
        this.CharacterFormatting = new CharacterFormatting();
        this.ParagraphFormatting = new ParagraphFormatting();
    }

    public Element(string text) :
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

    public Element AddElement()
    {
        var el = new Element();
        this.Add(el);
        return el;
    }

    public Element AddElement(string text)
    {
        var el = new Element(text);
        this.Add(el);
        return el;
    }

    public IEnumerable<Element> Elements
    {
        get { return this.Children.Where(n => n.NodeType == NodeType.Element).Cast<Element>(); }
    }
        
    internal MarkupRegions GetMarkupInfo()
    {
        var markupinfo = new MarkupRegions();

        int start_pos = 0;
        var region_stack = new Stack<Region>();

        foreach (var walkevent in this.Walk())
        {
            if (walkevent.Type == WalkEventType.EventEnter)
            {
                if (walkevent.Node is Element)
                {
                    var element = (Element) walkevent.Node;
                    var region = new Region(start_pos, element);
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
                        var field_region = new Region(start_pos,f);
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
                    throw new VisioOperationException(msg);
                }
            }
            else if (walkevent.Type == WalkEventType.EventExit)
            {
                if (walkevent.Node is Element)
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


        var writer = new VisioAutomation.ShapeSheet.Writers.SrcWriter();

        foreach (var region in regions_to_format)
        {

            // Apply character formatting
            var charcells = region.Element.CharacterFormatting;
            if (charcells != null)
            {
                var chars = shape.Characters;
                chars.Begin = region.Start;
                chars.End = region.End;
                chars.CharProps[ShapeSheet.SrcConstants.CharColor.Cell] = 0;
                short rownum = chars.CharPropsRow[(short) default_chars_bias];

                if (rownum < 0)
                {
                    throw new VisioAutomation.Exceptions.VisioOperationException("Could not create Character row");
                }

                writer.Clear();
                charcells.ApplyFormulas(writer, rownum);

                writer.Commit(shape, CellValueType.Formula);
            }

            // Apply paragraph formatting
            var paracells = region.Element.ParagraphFormatting;
            if (paracells != null)
            {
                var chars = shape.Characters;
                chars.Begin = region.Start;
                chars.End = region.End;
                chars.ParaProps[ShapeSheet.SrcConstants.ParaBullet.Cell] = 0;
                short rownum = chars.ParaPropsRow[(short) default_chars_bias];

                if (rownum < 0)
                {
                    throw new VisioAutomation.Exceptions.VisioOperationException("Could not create Paragraph row");
                }

                writer.Clear();
                paracells.ApplyFormulas(writer, rownum);

                writer.Commit(shape, CellValueType.Formula);
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