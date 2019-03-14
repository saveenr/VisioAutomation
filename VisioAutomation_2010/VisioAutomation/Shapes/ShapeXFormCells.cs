using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : CellGroup
    {
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral Angle { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.XFormPinX, this.PinX);
                yield return SrcValuePair.Create(SrcConstants.XFormPinY, this.PinY);
                yield return SrcValuePair.Create(SrcConstants.XFormLocPinX, this.LocPinX);
                yield return SrcValuePair.Create(SrcConstants.XFormLocPinY, this.LocPinY);
                yield return SrcValuePair.Create(SrcConstants.XFormWidth, this.Width);
                yield return SrcValuePair.Create(SrcConstants.XFormHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.XFormAngle, this.Angle);
            }
        }
    }
}