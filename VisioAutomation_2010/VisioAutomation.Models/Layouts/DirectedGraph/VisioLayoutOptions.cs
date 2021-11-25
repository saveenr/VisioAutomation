namespace VisioAutomation.Models.Layouts.DirectedGraph;

public class VisioLayoutOptions
{
    public VisioAutomation.Models.LayoutStyles.LayoutStyleBase VisioLayoutStyle;

    public VisioLayoutOptions()
    {
        var flowchart = new VisioAutomation.Models.LayoutStyles.FlowchartLayoutStyle();
        flowchart.LayoutDirection = LayoutStyles.LayoutDirection.TopToBottom;
        this.VisioLayoutStyle = flowchart;
    }
}