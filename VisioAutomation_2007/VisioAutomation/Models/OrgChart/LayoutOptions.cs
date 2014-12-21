using VA=VisioAutomation;

namespace VisioAutomation.Models.OrgChart
{
    public class LayoutOptions
    {
        public LayoutOptions()
        {
            DefaultNodeSize = new VA.Drawing.Size(2, 0.5);
            Direction = LayoutDirection.Down;
            UseDynamicConnectors = true;
        }

        public bool UseDynamicConnectors { get; set; }
        public LayoutDirection Direction { get; set; }
        public VA.Drawing.Size DefaultNodeSize { get; set; }
    }
}