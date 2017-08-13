using System.Collections.Generic;
using VisioAutomation.ShapeSheet;

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
            yield return new CellTuple(nameof(SrcConstants.LockAspect), SrcConstants.LockAspect, this.LockAspect);
            yield return new CellTuple(nameof(SrcConstants.LockBegin), SrcConstants.LockBegin, this.LockBegin);
            yield return new CellTuple(nameof(SrcConstants.LockCalcWH), SrcConstants.LockCalcWH, this.LockCalcWH);
            yield return new CellTuple(nameof(SrcConstants.LockCrop), SrcConstants.LockCrop, this.LockCrop);
            yield return new CellTuple(nameof(SrcConstants.LockCustomProp), SrcConstants.LockCustomProp, this.LockCustomProp);
            yield return new CellTuple(nameof(SrcConstants.LockDelete), SrcConstants.LockDelete, this.LockDelete);
            yield return new CellTuple(nameof(SrcConstants.LockEnd), SrcConstants.LockEnd, this.LockEnd);
            yield return new CellTuple(nameof(SrcConstants.LockFormat), SrcConstants.LockFormat, this.LockFormat);
            yield return new CellTuple(nameof(SrcConstants.LockFromGroupFormat), SrcConstants.LockFromGroupFormat, this.LockFromGroupFormat);
            yield return new CellTuple(nameof(SrcConstants.LockGroup), SrcConstants.LockGroup, this.LockGroup);
            yield return new CellTuple(nameof(SrcConstants.LockHeight), SrcConstants.LockHeight, this.LockHeight);
            yield return new CellTuple(nameof(SrcConstants.LockMoveX), SrcConstants.LockMoveX, this.LockMoveX);
            yield return new CellTuple(nameof(SrcConstants.LockMoveY), SrcConstants.LockMoveY, this.LockMoveY);
            yield return new CellTuple(nameof(SrcConstants.LockRotate), SrcConstants.LockRotate, this.LockRotate);
            yield return new CellTuple(nameof(SrcConstants.LockSelect), SrcConstants.LockSelect, this.LockSelect);
            yield return new CellTuple(nameof(SrcConstants.LockTextEdit), SrcConstants.LockTextEdit, this.LockTextEdit);
            yield return new CellTuple(nameof(SrcConstants.LockThemeColors), SrcConstants.LockThemeColors, this.LockThemeColors);
            yield return new CellTuple(nameof(SrcConstants.LockThemeEffects), SrcConstants.LockThemeEffects, this.LockThemeEffects);
            yield return new CellTuple(nameof(SrcConstants.LockVertexEdit), SrcConstants.LockVertexEdit, this.LockVertexEdit);
            yield return new CellTuple(nameof(SrcConstants.LockWidth), SrcConstants.LockWidth, this.LockWidth);
        }
    }
}