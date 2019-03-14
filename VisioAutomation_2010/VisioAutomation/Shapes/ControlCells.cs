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

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ControlCanGlue, this.CanGlue);
                yield return SrcValuePair.Create(SrcConstants.ControlTip, this.Tip);
                yield return SrcValuePair.Create(SrcConstants.ControlX, this.X);
                yield return SrcValuePair.Create(SrcConstants.ControlY, this.Y);
                yield return SrcValuePair.Create(SrcConstants.ControlYBehavior, this.YBehavior);
                yield return SrcValuePair.Create(SrcConstants.ControlXBehavior, this.XBehavior);
                yield return SrcValuePair.Create(SrcConstants.ControlXDynamics, this.XDynamics);
                yield return SrcValuePair.Create(SrcConstants.ControlYDynamics, this.YDynamics);
            }
        }


    }
}