using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.PageLayout
{
    public class CircularLayout : Layout
    {
        public CircularLayout() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Circular;
            this.ConnectorStyle = ConnectorStyle.CenterToCenter;
        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceCircular;
        }
    }
}