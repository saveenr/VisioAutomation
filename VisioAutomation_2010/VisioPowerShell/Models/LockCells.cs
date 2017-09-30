using System.Collections.Generic;
using SRCCON = VisioAutomation.ShapeSheet.SrcConstants;

namespace VisioPowerShell.Models
{
    public class LockCells : VisioPowerShell.Models.BaseCells
    {
        public string Aspect;
        public string Begin;
        public string CalcWH;
        public string Crop;
        public string CustomProp;
        public string Delete;
        public string End;
        public string Format;
        public string FromGroupFormat;
        public string Group;
        public string Height;
        public string MoveX;
        public string MoveY;
        public string Rotate;
        public string Select;
        public string TextEdit;
        public string ThemeColors;
        public string ThemeEffects;
        public string VertexEdit;
        public string Width;

        public override IEnumerable<CellTuple> GetCellTuples()
        {
            yield return new CellTuple(nameof(SRCCON.LockAspect), SRCCON.LockAspect, this.Aspect);
            yield return new CellTuple(nameof(SRCCON.LockBegin), SRCCON.LockBegin, this.Begin);
            yield return new CellTuple(nameof(SRCCON.LockCalcWH), SRCCON.LockCalcWH, this.CalcWH);
            yield return new CellTuple(nameof(SRCCON.LockCrop), SRCCON.LockCrop, this.Crop);
            yield return new CellTuple(nameof(SRCCON.LockCustomProp), SRCCON.LockCustomProp, this.CustomProp);
            yield return new CellTuple(nameof(SRCCON.LockDelete), SRCCON.LockDelete, this.Delete);
            yield return new CellTuple(nameof(SRCCON.LockEnd), SRCCON.LockEnd, this.End);
            yield return new CellTuple(nameof(SRCCON.LockFormat), SRCCON.LockFormat, this.Format);
            yield return new CellTuple(nameof(SRCCON.LockFromGroupFormat), SRCCON.LockFromGroupFormat, this.FromGroupFormat);
            yield return new CellTuple(nameof(SRCCON.LockGroup), SRCCON.LockGroup, this.Group);
            yield return new CellTuple(nameof(SRCCON.LockHeight), SRCCON.LockHeight, this.Height);
            yield return new CellTuple(nameof(SRCCON.LockMoveX), SRCCON.LockMoveX, this.MoveX);
            yield return new CellTuple(nameof(SRCCON.LockMoveY), SRCCON.LockMoveY, this.MoveY);
            yield return new CellTuple(nameof(SRCCON.LockRotate), SRCCON.LockRotate, this.Rotate);
            yield return new CellTuple(nameof(SRCCON.LockSelect), SRCCON.LockSelect, this.Select);
            yield return new CellTuple(nameof(SRCCON.LockTextEdit), SRCCON.LockTextEdit, this.TextEdit);
            yield return new CellTuple(nameof(SRCCON.LockThemeColors), SRCCON.LockThemeColors, this.ThemeColors);
            yield return new CellTuple(nameof(SRCCON.LockThemeEffects), SRCCON.LockThemeEffects, this.ThemeEffects);
            yield return new CellTuple(nameof(SRCCON.LockVertexEdit), SRCCON.LockVertexEdit, this.VertexEdit);
            yield return new CellTuple(nameof(SRCCON.LockWidth), SRCCON.LockWidth, this.Width);
        }
    }
}