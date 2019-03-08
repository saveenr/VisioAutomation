using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.LayoutStyles
{
    public class FlowchartLayoutStyle : LayoutStyleBase
    {
        public LayoutDirection LayoutDirection { get; set; }

        public FlowchartLayoutStyle()
        {
            this.ConnectorStyle = ConnectorStyle.Flowchart;
            this.LayoutDirection = LayoutDirection.TopToBottom;
        }

        protected override void SetPageCells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base.SetPageCells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) FlowchartLayoutStyle.GetPlaceStyle(this.LayoutDirection);
        }

        private static IVisio.VisCellVals GetPlaceStyle(LayoutDirection dir)
        {
            if (dir == LayoutDirection.TopToBottom)
            {
                return IVisio.VisCellVals.visPLOPlaceTopToBottom;
            }
            else if (dir == LayoutDirection.LeftToRight)
            {
                return IVisio.VisCellVals.visPLOPlaceLeftToRight;
            }
            else if (dir == LayoutDirection.BottomToTop)
            {
                return IVisio.VisCellVals.visPLOPlaceBottomToTop;
            }
            else if (dir == LayoutDirection.RightToLeft)
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
            return this.ConnectorsStyleAndDirectionToRouteStyle(this.ConnectorStyle, this.LayoutDirection);
        }
    }
}