namespace VisioAutomation.Layout.ContainerLayout
{
    public enum RenderStyle
    {
        UseVisioContainers,
        UseShapes
    }

    public class LayoutOptions
    {

        private double _itemWidth = 2.0;
        private double _containerHorizontalDistance = 1.0;
        private double _itemHeight = 0.25;
        private double _itemVerticalSpacing = 0.25/2.0;
        private double _padding = 0.25;
        private double _containerHeaderHeight = 0.5;

        public LayoutOptions()
        {
            Style = RenderStyle.UseVisioContainers;
        }

        public double ItemWidth
        {
            get { return _itemWidth; }
            set { _itemWidth = value; }
        }

        public double ContainerHorizontalDistance
        {
            get { return _containerHorizontalDistance; }
            set { _containerHorizontalDistance = value; }
        }

        public double ItemHeight
        {
            get { return _itemHeight; }
            set { _itemHeight = value; }
        }

        public double ItemVerticalSpacing
        {
            get { return _itemVerticalSpacing; }
            set { _itemVerticalSpacing = value; }
        }

        public double Padding
        {
            get { return _padding; }
            set { _padding = value; }
        }

        public RenderStyle Style { get; set; }

        public double ContainerHeaderHeight
        {
            get { return _containerHeaderHeight; }
            set { _containerHeaderHeight = value; }
        }
    }
}