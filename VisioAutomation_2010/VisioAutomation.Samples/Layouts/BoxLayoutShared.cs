using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using BoxL = VisioAutomation.Models.BoxLayout;

namespace VisioAutomationSamples
{
    public class BoxLayoutShared
    {
        public static void DrawNode(
            BoxL.Node node,
            VA.Drawing.Rectangle rect, IVisio.Page page)
        {           
            var shape = page.DrawRectangle(rect);
            node.Data = shape;
        }
    }
}