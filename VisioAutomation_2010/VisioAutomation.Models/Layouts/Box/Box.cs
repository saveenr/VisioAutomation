using System.Collections.Generic;

namespace VisioAutomation.Models.Layouts.Box;

public class Box : Node
{
    public Box(double w, double h) :
        this(new VisioAutomation.Geometry.Size(w, h) )
    {
    }

    protected Box(VisioAutomation.Geometry.Size s)
    {
        this.Size = s;
    }

    public override VisioAutomation.Geometry.Size CalculateSize()
    {
        return this.Size;
    }

    public override void _place(VisioAutomation.Geometry.Point origin)
    {
        this.Rectangle = new VisioAutomation.Geometry.Rectangle(origin, this.Size);
    }

    public override IEnumerable<Node> GetChildren()
    {
        yield break;
    }
}