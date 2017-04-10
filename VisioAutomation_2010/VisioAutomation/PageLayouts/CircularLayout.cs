using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.PageLayouts
{
    public class CircularLayout : LayoutBase
    {
        public CircularLayout()
        {
            this.LayoutStyle = LayoutStyle.Circular;
            this.ConnectorStyle = ConnectorStyle.CenterToCenter;
        }

        protected override void SetPageCells(PageLayoutFormulas pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceCircular;
        }
    }
}