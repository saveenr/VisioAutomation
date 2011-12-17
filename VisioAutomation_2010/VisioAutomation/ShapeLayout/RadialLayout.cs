using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.ShapeLayout
{
    public class RadialLayout : Layout
    {
        public RadialLayout() :
            base()
        {
            this.LayoutStyle = LayoutStyle.Radial;
            this.ConnectorStyle = ConnectorStyle.RightAngle;
        }

        public override void SetPageCells(VisioAutomation.Pages.PageCells pagecells)
        {
            base.SetPageCells(pagecells);
            pagecells.PlaceStyle = (int) IVisio.VisCellVals.visPLOPlaceDefault;
        }
    }
}