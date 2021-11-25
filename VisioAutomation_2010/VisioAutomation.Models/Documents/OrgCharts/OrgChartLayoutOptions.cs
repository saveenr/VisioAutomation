namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class OrgChartLayoutOptions
    {
        public bool UseDynamicConnectors;
        public OrgChartLayoutDirection Direction;
        public VisioAutomation.Core.Size DefaultNodeSize;
        public double PageBorderWidth;

        public OrgChartLayoutOptions()
        {
            this.DefaultNodeSize = new VisioAutomation.Core.Size(2, 0.5);
            this.Direction = OrgChartLayoutDirection.Down;
            this.UseDynamicConnectors = true;
            this.PageBorderWidth = 0.5;
        }
    }
}