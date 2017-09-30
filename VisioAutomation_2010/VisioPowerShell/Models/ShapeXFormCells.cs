using System.Collections.Generic;
using SRCCON=VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models
{
    public class ShapeXFormCells : VisioPowerShell.Models.BaseCells
    {
        public string Angle;
        public string Height;
        public string LocPinX;
        public string LocPinY;
        public string PinX;
        public string PinY;
        public string Width;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SRCCON.XFormAngle), SRCCON.XFormAngle, this.Angle);
            yield return new CellTuple(nameof(SRCCON.XFormHeight), SRCCON.XFormHeight, this.Height);
            yield return new CellTuple(nameof(SRCCON.XFormLocPinX), SRCCON.XFormLocPinX, this.LocPinX);
            yield return new CellTuple(nameof(SRCCON.XFormLocPinY), SRCCON.XFormLocPinY, this.LocPinY);
            yield return new CellTuple(nameof(SRCCON.XFormPinX), SRCCON.XFormPinX, this.PinX);
            yield return new CellTuple(nameof(SRCCON.XFormPinY), SRCCON.XFormPinY, this.PinY);
            yield return new CellTuple(nameof(SRCCON.XFormWidth), SRCCON.XFormWidth, this.Width);
        }
    }
}