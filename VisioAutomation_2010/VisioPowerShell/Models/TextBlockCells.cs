using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

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
            yield return new CellTuple(nameof(SrcConstants.TextXFormAngle), SrcConstants.TextXFormAngle, this.TextXFormAngle);
            yield return new CellTuple(nameof(SrcConstants.TextXFormHeight), SrcConstants.TextXFormHeight, this.TextXFormHeight);
            yield return new CellTuple(nameof(SrcConstants.TextXFormLocPinX), SrcConstants.TextXFormLocPinX, this.TextXFormLocPinX);
            yield return new CellTuple(nameof(SrcConstants.TextXFormLocPinY), SrcConstants.TextXFormLocPinY, this.TextXFormLocPinY);
            yield return new CellTuple(nameof(SrcConstants.TextXFormPinX), SrcConstants.TextXFormPinX, this.TextXFormPinX);
            yield return new CellTuple(nameof(SrcConstants.TextXFormPinY), SrcConstants.TextXFormPinY, this.TextXFormPinY);
            yield return new CellTuple(nameof(SrcConstants.TextXFormWidth), SrcConstants.TextXFormWidth, this.TextXFormWidth);
        }
    }
}