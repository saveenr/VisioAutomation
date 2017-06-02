namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class LayoutOptions
    {
        public LayoutOptions()
        {
            this.DefaultNodeSize = new Geometry.Size(2, 0.5);
            this.Direction = LayoutDirection.Down;
            this.UseDynamicConnectors = true;
        }

        public bool UseDynamicConnectors { get; set; }
        public LayoutDirection Direction { get; set; }
        public Geometry.Size DefaultNodeSize { get; set; }
    }
}