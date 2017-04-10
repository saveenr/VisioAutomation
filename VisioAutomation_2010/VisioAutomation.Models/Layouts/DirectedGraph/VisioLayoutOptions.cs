using VisioAutomation.PageLayouts;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class VisioLayoutOptions
    {
        public LayoutBase Layout;

        public VisioLayoutOptions()
        {
            var flowchart = new FlowchartLayout();
            flowchart.LayoutDirection = PageLayouts.LayoutDirection.TopToBottom;
            this.Layout = flowchart;
        }        
    }
}