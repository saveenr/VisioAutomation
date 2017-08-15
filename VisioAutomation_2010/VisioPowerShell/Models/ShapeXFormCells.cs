using System.Collections.Generic;
using SRCCON=VisioAutomation.ShapeSheet.SrcConstants;

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
            yield return new CellTuple(nameof(SRCCON.XFormAngle), SRCCON.XFormAngle, this.XFormAngle);
            yield return new CellTuple(nameof(SRCCON.XFormHeight), SRCCON.XFormHeight, this.XFormHeight);
            yield return new CellTuple(nameof(SRCCON.XFormLocPinX), SRCCON.XFormLocPinX, this.XFormLocPinX);
            yield return new CellTuple(nameof(SRCCON.XFormLocPinY), SRCCON.XFormLocPinY, this.XFormLocPinY);
            yield return new CellTuple(nameof(SRCCON.XFormPinX), SRCCON.XFormPinX, this.XFormPinX);
            yield return new CellTuple(nameof(SRCCON.XFormPinY), SRCCON.XFormPinY, this.XFormPinY);
            yield return new CellTuple(nameof(SRCCON.XFormWidth), SRCCON.XFormWidth, this.XFormWidth);
        }
    }
}