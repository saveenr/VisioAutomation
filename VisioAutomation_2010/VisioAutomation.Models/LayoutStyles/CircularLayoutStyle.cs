using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.LayoutStyles
{
    public class CircularLayoutStyle : LayoutStyleBase
    {
        public CircularLayoutStyle()
        {
            this.ConnectorStyle = ConnectorStyle.CenterToCenter;
        }

        protected override void SetPageCells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base.SetPageCells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceCircular;
        }
    }
}