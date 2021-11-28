using VisioAutomation.Extensions;
using VisioAutomation.Models.Layouts.Box;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;

namespace VSamples
{
    public class BoxLayoutShared
    {
        public static void DrawNode(
            Node node,
            VA.Core.Rectangle rect, IVisio.Page page)
        {           
            var shape = page.DrawRectangle(rect);
            node.Data = shape;
        }
    }
}