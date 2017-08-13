using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

namespace VisioPowerShell.Models
{
    public class ShapeXFormCells : VisioPowerShell.Models.BaseCells
    {
        public string XFormAngle;
        public string XFormHeight;
        public string XFormLocPinX;
        public string XFormLocPinY;
        public string XFormPinX;
        public string XFormPinY;
        public string XFormWidth;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SrcConstants.XFormAngle), SrcConstants.XFormAngle, this.XFormAngle);
            yield return new CellTuple(nameof(SrcConstants.XFormHeight), SrcConstants.XFormHeight, this.XFormHeight);
            yield return new CellTuple(nameof(SrcConstants.XFormLocPinX), SrcConstants.XFormLocPinX, this.XFormLocPinX);
            yield return new CellTuple(nameof(SrcConstants.XFormLocPinY), SrcConstants.XFormLocPinY, this.XFormLocPinY);
            yield return new CellTuple(nameof(SrcConstants.XFormPinX), SrcConstants.XFormPinX, this.XFormPinX);
            yield return new CellTuple(nameof(SrcConstants.XFormPinY), SrcConstants.XFormPinY, this.XFormPinY);
            yield return new CellTuple(nameof(SrcConstants.XFormWidth), SrcConstants.XFormWidth, this.XFormWidth);
        }
    }
}