using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages.PageLayout
{
    public class RadialLayout : Layout
    {
        public RadialLayout() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Radial;
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        protected override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }
}