namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class LayoutOptions
    {
        public bool UseDynamicConnectors;
        public LayoutDirection Direction;
        public VisioAutomation.Geometry.Size DefaultNodeSize;
        public double PageBorderWidth;

        public LayoutOptions()
        {
            this.DefaultNodeSize = new VisioAutomation.Geometry.Size(2, 0.5);
            this.Direction = LayoutDirection.Down;
            this.UseDynamicConnectors = true;
            this.PageBorderWidth = 0.5;
        }
    }
}