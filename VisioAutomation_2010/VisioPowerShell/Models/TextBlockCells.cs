using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models
{
    public class TextBlockCells : VisioPowerShell.Models.BaseCells
    {
        public string TextXFormAngle;
        public string TextXFormHeight;
        public string TextXFormLocPinX;
        public string TextXFormLocPinY;
        public string TextXFormPinX;
        public string TextXFormPinY;
        public string TextXFormWidth;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SRCCON.TextXFormAngle), SRCCON.TextXFormAngle, this.TextXFormAngle);
            yield return new CellTuple(nameof(SRCCON.TextXFormHeight), SRCCON.TextXFormHeight, this.TextXFormHeight);
            yield return new CellTuple(nameof(SRCCON.TextXFormLocPinX), SRCCON.TextXFormLocPinX, this.TextXFormLocPinX);
            yield return new CellTuple(nameof(SRCCON.TextXFormLocPinY), SRCCON.TextXFormLocPinY, this.TextXFormLocPinY);
            yield return new CellTuple(nameof(SRCCON.TextXFormPinX), SRCCON.TextXFormPinX, this.TextXFormPinX);
            yield return new CellTuple(nameof(SRCCON.TextXFormPinY), SRCCON.TextXFormPinY, this.TextXFormPinY);
            yield return new CellTuple(nameof(SRCCON.TextXFormWidth), SRCCON.TextXFormWidth, this.TextXFormWidth);
        }
    }
}