using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models
{
    public class LockCells : VisioPowerShell.Models.BaseCells
    {
        public string LockAspect;
        public string LockBegin;
        public string LockCalcWH;
        public string LockCrop;
        public string LockCustomProp;
        public string LockDelete;
        public string LockEnd;
        public string LockFormat;
        public string LockFromGroupFormat;
        public string LockGroup;
        public string LockHeight;
        public string LockMoveX;
        public string LockMoveY;
        public string LockRotate;
        public string LockSelect;
        public string LockTextEdit;
        public string LockThemeColors;
        public string LockThemeEffects;
        public string LockVertexEdit;
        public string LockWidth;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SRCCON.LockAspect), SRCCON.LockAspect, this.LockAspect);
            yield return new CellTuple(nameof(SRCCON.LockBegin), SRCCON.LockBegin, this.LockBegin);
            yield return new CellTuple(nameof(SRCCON.LockCalcWH), SRCCON.LockCalcWH, this.LockCalcWH);
            yield return new CellTuple(nameof(SRCCON.LockCrop), SRCCON.LockCrop, this.LockCrop);
            yield return new CellTuple(nameof(SRCCON.LockCustomProp), SRCCON.LockCustomProp, this.LockCustomProp);
            yield return new CellTuple(nameof(SRCCON.LockDelete), SRCCON.LockDelete, this.LockDelete);
            yield return new CellTuple(nameof(SRCCON.LockEnd), SRCCON.LockEnd, this.LockEnd);
            yield return new CellTuple(nameof(SRCCON.LockFormat), SRCCON.LockFormat, this.LockFormat);
            yield return new CellTuple(nameof(SRCCON.LockFromGroupFormat), SRCCON.LockFromGroupFormat, this.LockFromGroupFormat);
            yield return new CellTuple(nameof(SRCCON.LockGroup), SRCCON.LockGroup, this.LockGroup);
            yield return new CellTuple(nameof(SRCCON.LockHeight), SRCCON.LockHeight, this.LockHeight);
            yield return new CellTuple(nameof(SRCCON.LockMoveX), SRCCON.LockMoveX, this.LockMoveX);
            yield return new CellTuple(nameof(SRCCON.LockMoveY), SRCCON.LockMoveY, this.LockMoveY);
            yield return new CellTuple(nameof(SRCCON.LockRotate), SRCCON.LockRotate, this.LockRotate);
            yield return new CellTuple(nameof(SRCCON.LockSelect), SRCCON.LockSelect, this.LockSelect);
            yield return new CellTuple(nameof(SRCCON.LockTextEdit), SRCCON.LockTextEdit, this.LockTextEdit);
            yield return new CellTuple(nameof(SRCCON.LockThemeColors), SRCCON.LockThemeColors, this.LockThemeColors);
            yield return new CellTuple(nameof(SRCCON.LockThemeEffects), SRCCON.LockThemeEffects, this.LockThemeEffects);
            yield return new CellTuple(nameof(SRCCON.LockVertexEdit), SRCCON.LockVertexEdit, this.LockVertexEdit);
            yield return new CellTuple(nameof(SRCCON.LockWidth), SRCCON.LockWidth, this.LockWidth);
        }
    }
}