namespace VisioAutomation.Models.Documents.OrgCharts
{
    public class LayoutOptions
    {
        public LayoutOptions()
        {
            this.DefaultNodeSize = new Drawing.Size(2, 0.5);
            this.Direction = LayoutDirection.Down;
            this.UseDynamicConnectors = true;
        }

        public bool UseDynamicConnectors { get; set; }
        public LayoutDirection Direction { get; set; }
        public Drawing.Size DefaultNodeSize { get; set; }
    }
}