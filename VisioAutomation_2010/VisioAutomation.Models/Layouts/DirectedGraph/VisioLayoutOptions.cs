namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class VisioLayoutOptions
    {
        public Pages.PageLayout.Layout Layout;

        public VisioLayoutOptions()
        {
            var flowchart = new Pages.PageLayout.FlowchartLayout();
            flowchart.LayoutDirection = Pages.PageLayout.LayoutDirection.TopToBottom;
            this.Layout = flowchart;
        }        
    }
}