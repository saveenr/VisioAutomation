using VisioAutomation.Models.LayoutStyles;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class VisioLayoutOptions
    {
        public LayoutStyleBase Layout;

        public VisioLayoutOptions()
        {
            var flowchart = new FlowchartLayoutStyle();
            flowchart.LayoutDirection = LayoutStyles.LayoutDirection.TopToBottom;
            this.Layout = flowchart;
        }        
    }
}