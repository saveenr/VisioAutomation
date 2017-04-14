using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.PageLayouts
{
    public class CompactTreeLayout : LayoutBase
    {
        public CompactTreeDirection Direction { get; set; }

        public CompactTreeLayout()
        {
            this.LayoutStyle = LayoutStyle.CompactTree;
            this.ConnectorStyle = ConnectorStyle.OrganizationChart;
            this.Direction = CompactTreeDirection.DownThenRight;
        }

        protected override void SetPageCells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base.SetPageCells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) CompactTreeLayout.GetPlaceStyle(this.Direction);
        }

        private static IVisio.VisCellVals GetPlaceStyle(CompactTreeDirection dir)
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