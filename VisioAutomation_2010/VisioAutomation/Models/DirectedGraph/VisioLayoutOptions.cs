using VA=VisioAutomation;

namespace VisioAutomation.Models.DirectedGraph
{
    public class VisioLayoutOptions
    {
        public VA.Pages.PageLayout.Layout Layout;

        public VisioLayoutOptions()
        {
            var flowchart = new VA.Pages.PageLayout.FlowchartLayout();
            flowchart.Direction = VA.Pages.PageLayout.Direction.TopToBottom;
            this.Layout = flowchart;
        }        
    }
}