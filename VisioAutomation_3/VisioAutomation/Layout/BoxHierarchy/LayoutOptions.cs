using VA=VisioAutomation;

namespace VisioAutomation
{
}

namespace VisioAutomation.Layout.BoxHierarchy
{
    public class LayoutOptions
    {
        private VA.DirectionVertical _directionVertical = VA.DirectionVertical.Up;
        private VA.DirectionHorizontal _directionHorizontal = VA.DirectionHorizontal.Right;
        private double _defaultWidth = 1.0;
        private double _defaultHeight = 1.0;

        public VA.DirectionVertical DirectionVertical
        {
            get { return _directionVertical; }
            set { _directionVertical = value; }
        }

        public VA.DirectionHorizontal DirectionHorizontal
        {
            get { return _directionHorizontal; }
            set { _directionHorizontal = value; }
        }

        public double DefaultWidth
        {
            get { return _defaultWidth; }
            set { _defaultWidth = value; }
        }

        public double DefaultHeight
        {
            get { return _defaultHeight; }
            set { _defaultHeight = value; }
        }
    }
}