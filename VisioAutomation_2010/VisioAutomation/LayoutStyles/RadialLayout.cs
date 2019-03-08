using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.LayoutStyles
{
    public class RadialLayoutStyle : LayoutStyleBase
    {
        public RadialLayoutStyle()
        {
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        protected override void SetPageCells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base.SetPageCells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }
}