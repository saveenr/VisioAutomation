using VA=VisioAutomation;

namespace VisioAutomation.Layout.BoxLayout
{
    public class LayoutOptions
    {
        public VA.Drawing.Point Origin { get; set; }
        public DirectionVertical DirectionVertical { get; set; }
        public DirectionHorizontal DirectionHorizontal { get; set; }
        public double DefaultWidth { get; set; }
        public double DefaultHeight { get; set; }

        public LayoutOptions()
        {
            this.DefaultHeight = 1.0;
            this.DefaultWidth = 1.0;
            this.DirectionHorizontal = DirectionHorizontal.LeftToRight;
            this.DirectionVertical = DirectionVertical.BottomToTop;
            this.Origin = new VA.Drawing.Point(0, 0);
        }
    }
}