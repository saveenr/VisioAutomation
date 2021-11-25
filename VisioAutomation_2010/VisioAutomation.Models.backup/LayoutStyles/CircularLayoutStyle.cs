using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.LayoutStyles
{
    public class CircularLayoutStyle : LayoutStyleBase
    {
        public CircularLayoutStyle()
        {
            this.ConnectorStyle = ConnectorStyle.CenterToCenter;
        }

        protected override void _set_page_cells(VisioAutomation.Pages.PageLayoutCells page_layout_cells)
        {
            base._set_page_cells(page_layout_cells);
            page_layout_cells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceCircular;
        }
    }
}