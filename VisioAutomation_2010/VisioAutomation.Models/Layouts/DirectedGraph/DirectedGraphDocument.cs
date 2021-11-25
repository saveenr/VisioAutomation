

namespace VisioAutomation.Models.Layouts.DirectedGraph;

public class DirectedGraphDocument
{
    public readonly List<DirectedGraphLayout> Layouts;
    public string Template;
    public VisioAutomation.Geometry.Size BorderSize;

    public DirectedGraphDocument()
    {
        this.Layouts = new List<DirectedGraphLayout>();
        this.Template = null;
        this.BorderSize = new VisioAutomation.Geometry.Size(1.0, 1.0);
    }
}