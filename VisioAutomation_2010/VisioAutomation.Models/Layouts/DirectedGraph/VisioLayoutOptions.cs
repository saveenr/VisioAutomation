namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class VisioLayoutOptions
    {
        public VisioAutomation.Models.LayoutStyles.LayoutStyleBase Layout;

        public string EdgeMasterName = "Dynamic Connector";
        public string EdgeStencilName = "connec_u.vss";

        public VisioLayoutOptions()
        {
            var flowchart = new VisioAutomation.Models.LayoutStyles.FlowchartLayoutStyle();
            flowchart.LayoutDirection = LayoutStyles.LayoutDirection.TopToBottom;
            this.Layout = flowchart;
        }        
    }
}