namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    public class DirectedGraphStyling
    {
        public VisioAutomation.Models.LayoutStyles.LayoutStyleBase VisioLayoutStyle;

        public string EdgeMasterName = "Dynamic Connector";
        public string EdgeStencilName = "connec_u.vss";

        public DirectedGraphStyling()
        {
            var flowchart = new VisioAutomation.Models.LayoutStyles.FlowchartLayoutStyle();
            flowchart.LayoutDirection = LayoutStyles.LayoutDirection.TopToBottom;
            this.VisioLayoutStyle = flowchart;
        }        
    }
}