

namespace VisioAutomation.Models.LayoutStyles
{
    public class CompactTreeLayout : LayoutStyleBase
    {
        public CompactTreeDirection Direction { get; set; }

        public CompactTreeLayout()
        {
            this.ConnectorStyle = ConnectorStyle.OrganizationChart;
            this.Direction = CompactTreeDirection.DownThenRight;
        }

        protected override void _set_page_cells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base._set_page_cells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) CompactTreeLayout._get_place_style(this.Direction);
        }

        private static IVisio.VisCellVals _get_place_style(CompactTreeDirection dir)
        {
            if (dir == CompactTreeDirection.DownThenRight)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactDownRight;
            }
            else if (dir == CompactTreeDirection.RightThenDown)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactRightDown;
            }
            else if (dir == CompactTreeDirection.RightThenUp)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactRightUp;
            }
            else if (dir == CompactTreeDirection.UpThenRigtht)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactUpRight;
            }
            else if (dir == CompactTreeDirection.UpThenLeft)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactUpLeft;
            }
            else if (dir == CompactTreeDirection.LeftThenUp)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactLeftUp;
            }
            else if (dir == CompactTreeDirection.LeftThenDown)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactLeftDown;
            }
            else if (dir == CompactTreeDirection.DownThenLeft)
            {
                return IVisio.VisCellVals.visPLOPlaceCompactDownLeft;
            }
            else
            {
                throw new System.ArgumentException(nameof(dir));
            }
        }
    }
}