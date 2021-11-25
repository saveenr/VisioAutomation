

namespace VisioAutomation.Models.Text;

public class FieldBase : Node
{
    private const string Placeholder = "[FIELD]";
    public IVisio.VisFieldFormats Format { get; set; }

    internal FieldBase(VisioAutomation.Models.Text.NodeType nt)
        : base(nt)
    {
    }
        
    public string PlaceholderText
    {
        get
        {
            return FieldBase.Placeholder;
        }
    }
}