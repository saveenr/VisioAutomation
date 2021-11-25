

namespace VisioAutomation.Models.LayoutStyles
{
    public class RadialLayoutStyle : LayoutStyleBase
    {
        public RadialLayoutStyle()
        {
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        protected override void _set_page_cells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base._set_page_cells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }
}