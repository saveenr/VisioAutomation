using IVisio = Microsoft.Office.Interop.Visio;
using VA=VisioAutomation;

namespace VisioAutomation.Controls
{
    public class ControlQuery : VA.ShapeSheet.Query.SectionQuery
    {
        public VA.ShapeSheet.Query.SectionQueryColumn CanGlue { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Tip { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn X { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn Y { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn YBehavior { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn XBehavior { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn XDynamics { get; set; }
        public VA.ShapeSheet.Query.SectionQueryColumn YDynamics { get; set; }

        public ControlQuery() :
            base(IVisio.VisSectionIndices.visSectionControls)
        {
            this.CanGlue = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_CanGlue.Cell, "CanGlue");
            this.Tip = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_Tip.Cell, "Tip");
            this.X = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_X.Cell, "X");
            this.Y = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_Y.Cell, "Y");
            this.YBehavior = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_YCon.Cell, "YBehavior");
            this.XBehavior = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_XCon.Cell, "XBehavior");
            this.XDynamics = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_XDyn.Cell, "XDynamics");
            this.YDynamics = this.AddColumn(VA.ShapeSheet.SRCConstants.Controls_YDyn.Cell, "YDynamics");
        }

    }
}