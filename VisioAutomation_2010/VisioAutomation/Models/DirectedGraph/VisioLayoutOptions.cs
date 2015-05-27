namespace VisioAutomation.Models.DirectedGraph
{
    public class VisioLayoutOptions
    {
        public Pages.PageLayout.Layout Layout;

        public VisioLayoutOptions()
        {
            var flowchart = new Pages.PageLayout.FlowchartLayout();
            flowchart.Direction = Pages.PageLayout.Direction.TopToBottom;
            this.Layout = flowchart;
        }        
    }
}