using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models
{
    public class TextXFormCells : VisioPowerShell.Models.BaseCells
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
            yield return new CellTuple(nameof(SRCCON.TextXFormAngle), SRCCON.TextXFormAngle, this.Angle);
            yield return new CellTuple(nameof(SRCCON.TextXFormHeight), SRCCON.TextXFormHeight, this.Height);
            yield return new CellTuple(nameof(SRCCON.TextXFormLocPinX), SRCCON.TextXFormLocPinX, this.LocPinX);
            yield return new CellTuple(nameof(SRCCON.TextXFormLocPinY), SRCCON.TextXFormLocPinY, this.LocPinY);
            yield return new CellTuple(nameof(SRCCON.TextXFormPinX), SRCCON.TextXFormPinX, this.PinX);
            yield return new CellTuple(nameof(SRCCON.TextXFormPinY), SRCCON.TextXFormPinY, this.PinY);
            yield return new CellTuple(nameof(SRCCON.TextXFormWidth), SRCCON.TextXFormWidth, this.Width);
        }
    }
}