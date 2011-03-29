using VisioAutomation.Layout.MSAGL;
using VA=VisioAutomation;

namespace VisioAutomation.Scripting.FlowChart
{
    public class RenderItem
    {
        public Layout.MSAGL.Drawing Drawing { get; set; }
        public DirectedGraphLayout DirectedGraphLayout { get; set; }

        public RenderItem(Layout.MSAGL.Drawing drawing, VA.Layout.MSAGL.DirectedGraphLayout layout)
        {
            this.Drawing = drawing;
            this.DirectedGraphLayout = layout;
        }
    }
}