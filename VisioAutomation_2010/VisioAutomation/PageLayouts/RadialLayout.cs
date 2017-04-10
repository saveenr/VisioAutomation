using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.PageLayouts
{
    public class RadialLayout : LayoutBase
    {
        public RadialLayout()
        {
            this.LayoutStyle = LayoutStyle.Radial;
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        protected override void SetPageCells(PageLayoutFormulas pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }
}