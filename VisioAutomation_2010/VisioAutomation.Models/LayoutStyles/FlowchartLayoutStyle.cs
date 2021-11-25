

namespace VisioAutomation.Models.LayoutStyles
{
    public class FlowchartLayoutStyle : LayoutStyleBase
    {
        public LayoutDirection LayoutDirection { get; set; }

        public FlowchartLayoutStyle()
        {
            this.ConnectorStyle = ConnectorStyle.Flowchart;
            this.LayoutDirection = LayoutDirection.TopToBottom;
        }

        protected override void _set_page_cells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base._set_page_cells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) FlowchartLayoutStyle._get_place_style(this.LayoutDirection);
        }

        private static IVisio.VisCellVals _get_place_style(LayoutDirection dir)
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