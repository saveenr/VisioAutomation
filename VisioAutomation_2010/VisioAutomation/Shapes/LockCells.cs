using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class LockCells : CellGroup
    {
        public CellValueLiteral Aspect { get; set; }
        public CellValueLiteral Begin { get; set; }
        public CellValueLiteral CalcWH { get; set; }
        public CellValueLiteral Crop { get; set; }
        public CellValueLiteral CustProp { get; set; }
        public CellValueLiteral Delete { get; set; }
        public CellValueLiteral End { get; set; }
        public CellValueLiteral Format { get; set; }
        public CellValueLiteral FromGroupFormat { get; set; }
        public CellValueLiteral Group { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral MoveX { get; set; }
        public CellValueLiteral MoveY { get; set; }
        public CellValueLiteral Rotate { get; set; }
        public CellValueLiteral Select { get; set; }
        public CellValueLiteral TextEdit { get; set; }
        public CellValueLiteral ThemeColors { get; set; }
        public CellValueLiteral ThemeEffects { get; set; }
        public CellValueLiteral VertexEdit { get; set; }
        public CellValueLiteral Width { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.LockAspect, this.Aspect);
                yield return SrcValuePair.Create(SrcConstants.LockBegin, this.Begin);
                yield return SrcValuePair.Create(SrcConstants.LockCalcWH, this.CalcWH);
                yield return SrcValuePair.Create(SrcConstants.LockCrop, this.Crop);
                yield return SrcValuePair.Create(SrcConstants.LockCustomProp, this.CustProp);
                yield return SrcValuePair.Create(SrcConstants.LockDelete, this.Delete);
                yield return SrcValuePair.Create(SrcConstants.LockEnd, this.End);
                yield return SrcValuePair.Create(SrcConstants.LockFormat, this.Format);
                yield return SrcValuePair.Create(SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
                yield return SrcValuePair.Create(SrcConstants.LockGroup, this.Group);
                yield return SrcValuePair.Create(SrcConstants.LockHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.LockMoveX, this.MoveX);
                yield return SrcValuePair.Create(SrcConstants.LockMoveY, this.MoveY);
                yield return SrcValuePair.Create(SrcConstants.LockRotate, this.Rotate);
                yield return SrcValuePair.Create(SrcConstants.LockSelect, this.Select);
                yield return SrcValuePair.Create(SrcConstants.LockTextEdit, this.TextEdit);
                yield return SrcValuePair.Create(SrcConstants.LockThemeColors, this.ThemeColors);
                yield return SrcValuePair.Create(SrcConstants.LockThemeEffects, this.ThemeEffects);
                yield return SrcValuePair.Create(SrcConstants.LockVertexEdit, this.VertexEdit);
                yield return SrcValuePair.Create(SrcConstants.LockWidth, this.Width);
            }
        }


    }
}