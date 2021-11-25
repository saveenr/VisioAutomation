using VisioAutomation.Models.Layouts.Box;
using VisioAutomation.Extensions;

namespace VisioAutomationSamples;

public class BoxLayoutShared
{
    public static void DrawNode(
        Node node,
        VA.Geometry.Rectangle rect, IVisio.Page page)
    {           
        var shape = page.DrawRectangle(rect);
        node.Data = shape;
    }
}