namespace VisioAutomation.Models.ContainerLayout
{
    public class LayoutOptions
    {
        public string ManualItemStencil = "basic_u.vss";
        public string ManualItemMaster = "Rounded Rectangle";
        public string ManualContainerMaster = "Rectangle";
        public string ContainerMaster = "Container 1";

        public double ItemWidth { get; set; }
        public double ContainerHorizontalDistance { get; set; }
        public double ItemHeight { get; set; }
        public double ItemVerticalSpacing { get; set; }
        public double Padding { get; set; }
        public double ContainerHeaderHeight { get; set; }

        public Formatting ContainerFormatting { get; set; }
        public Formatting ContainerItemFormatting { get; set; }

        public LayoutOptions()
        {
            ContainerHeaderHeight = 0.25;
            Padding = 0.125;
            ItemVerticalSpacing = 0.125;
            ItemHeight = 0.25;
            ContainerHorizontalDistance = 1.0;
            ItemWidth = 2.0;
            this.ContainerFormatting = new Formatting();
            this.ContainerItemFormatting = new Formatting();
            this.ContainerFormatting.TextCells.VerticalAlign = "0";
        }
    }
}