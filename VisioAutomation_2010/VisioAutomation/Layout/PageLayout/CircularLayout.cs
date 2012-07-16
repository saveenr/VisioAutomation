using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Layout.PageLayout
{
    public class CircularLayout : Layout
    {
        public CircularLayout() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Circular;
            this.ConnectorStyle = ConnectorStyle.CenterToCenter;
        }

        protected override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceCircular;
        }
    }
}