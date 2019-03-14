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

        public override IEnumerable<NamedSrcValuePair> NamedSrcValuePairs
        {
            get
            {


                yield return NamedSrcValuePair.Create(nameof(this.Aspect), SrcConstants.LockAspect, this.Aspect);
                yield return NamedSrcValuePair.Create(nameof(this.Begin), SrcConstants.LockBegin, this.Begin);
                yield return NamedSrcValuePair.Create(nameof(this.CalcWH), SrcConstants.LockCalcWH, this.CalcWH);
                yield return NamedSrcValuePair.Create(nameof(this.Crop), SrcConstants.LockCrop, this.Crop);
                yield return NamedSrcValuePair.Create(nameof(this.CustProp), SrcConstants.LockCustomProp, this.CustProp);
                yield return NamedSrcValuePair.Create(nameof(this.Delete), SrcConstants.LockDelete, this.Delete);
                yield return NamedSrcValuePair.Create(nameof(this.End), SrcConstants.LockEnd, this.End);
                yield return NamedSrcValuePair.Create(nameof(this.Format), SrcConstants.LockFormat, this.Format);
                yield return NamedSrcValuePair.Create(nameof(this.FromGroupFormat), SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
                yield return NamedSrcValuePair.Create(nameof(this.Group), SrcConstants.LockGroup, this.Group);
                yield return NamedSrcValuePair.Create(nameof(this.Height), SrcConstants.LockHeight, this.Height);
                yield return NamedSrcValuePair.Create(nameof(this.MoveX), SrcConstants.LockMoveX, this.MoveX);
                yield return NamedSrcValuePair.Create(nameof(this.MoveY), SrcConstants.LockMoveY, this.MoveY);
                yield return NamedSrcValuePair.Create(nameof(this.Rotate), SrcConstants.LockRotate, this.Rotate);
                yield return NamedSrcValuePair.Create(nameof(this.Select), SrcConstants.LockSelect, this.Select);
                yield return NamedSrcValuePair.Create(nameof(this.TextEdit), SrcConstants.LockTextEdit, this.TextEdit);
                yield return NamedSrcValuePair.Create(nameof(this.ThemeColors), SrcConstants.LockThemeColors, this.ThemeColors);
                yield return NamedSrcValuePair.Create(nameof(this.ThemeEffects), SrcConstants.LockThemeEffects, this.ThemeEffects);
                yield return NamedSrcValuePair.Create(nameof(this.VertexEdit), SrcConstants.LockVertexEdit, this.VertexEdit);
                yield return NamedSrcValuePair.Create(nameof(this.Width), SrcConstants.LockWidth, this.Width);
            }
        }


    }
}