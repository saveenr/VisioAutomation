using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ControlCells : CellGroup
    {
        public CellValueLiteral CanGlue { get; set; }
        public CellValueLiteral Tip { get; set; }
        public CellValueLiteral X { get; set; }
        public CellValueLiteral Y { get; set; }
        public CellValueLiteral YBehavior { get; set; }
        public CellValueLiteral XBehavior { get; set; }
        public CellValueLiteral XDynamics { get; set; }
        public CellValueLiteral YDynamics { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {
                yield return CellMetadataItem.Create(nameof(this.CanGlue), SrcConstants.ControlCanGlue, this.CanGlue);
                yield return CellMetadataItem.Create(nameof(this.Tip), SrcConstants.ControlTip, this.Tip);
                yield return CellMetadataItem.Create(nameof(this.X), SrcConstants.ControlX, this.X);
                yield return CellMetadataItem.Create(nameof(this.Y), SrcConstants.ControlY, this.Y);
                yield return CellMetadataItem.Create(nameof(this.YBehavior), SrcConstants.ControlYBehavior, this.YBehavior);
                yield return CellMetadataItem.Create(nameof(this.XBehavior), SrcConstants.ControlXBehavior, this.XBehavior);
                yield return CellMetadataItem.Create(nameof(this.XDynamics), SrcConstants.ControlXDynamics, this.XDynamics);
                yield return CellMetadataItem.Create(nameof(this.YDynamics), SrcConstants.ControlYDynamics, this.YDynamics);
            }
        }


    }
}