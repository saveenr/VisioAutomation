using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages.PageLayout
{
    public class FlowchartLayout : Layout
    {
        public Direction Direction { get; set; }

        public FlowchartLayout()
        {
            this.LayoutStyle = LayoutStyle.Flowchart;
            this.ConnectorStyle = ConnectorStyle.Flowchart;
            this.Direction = Direction.TopToBottom;
        }

        protected override void SetPageCells(PageLayoutFormulas pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) FlowchartLayout.GetPlaceStyle(this.Direction);
        }

        private static IVisio.VisCellVals GetPlaceStyle(Direction dir)
        {
            if (dir == Direction.TopToBottom)
            {
                return IVisio.VisCellVals.visPLOPlaceTopToBottom;
            }
            else if (dir == Direction.LeftToRight)
            {
                return IVisio.VisCellVals.visPLOPlaceLeftToRight;
            }
            else if (dir == Direction.BottomToTop)
            {
                return IVisio.VisCellVals.visPLOPlaceBottomToTop;
            }
            else if (dir == Direction.RightToLeft)
            {
                return IVisio.VisCellVals.visPLOPlaceRightToLeft;
            }
            else
            {
                throw new System.ArgumentException(nameof(dir));
            }
        }

        protected override IVisio.VisCellVals? ConnectorsStyleToRouteStyle()
        {
            var rs = base.ConnectorsStyleToRouteStyle();
            if (rs.HasValue)
            {
                return rs;
            }
            return this.ConnectorsStyleAndDirectionToRouteStyle(this.ConnectorStyle, this.Direction);
        }
    }
}