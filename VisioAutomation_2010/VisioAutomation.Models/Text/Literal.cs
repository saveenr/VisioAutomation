namespace VisioAutomation.Models.Text;

public class Literal : Node
{
    public Literal(string text) : 
        base(NodeType.Literal)
    {
        this.Text = text;
    }

    public string Text { get; set; }
}