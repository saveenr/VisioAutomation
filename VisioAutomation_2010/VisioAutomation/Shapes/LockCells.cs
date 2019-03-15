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

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.Aspect), SrcConstants.LockAspect, this.Aspect);
                yield return CellMetadataItem.Create(nameof(this.Begin), SrcConstants.LockBegin, this.Begin);
                yield return CellMetadataItem.Create(nameof(this.CalcWH), SrcConstants.LockCalcWH, this.CalcWH);
                yield return CellMetadataItem.Create(nameof(this.Crop), SrcConstants.LockCrop, this.Crop);
                yield return CellMetadataItem.Create(nameof(this.CustProp), SrcConstants.LockCustomProp, this.CustProp);
                yield return CellMetadataItem.Create(nameof(this.Delete), SrcConstants.LockDelete, this.Delete);
                yield return CellMetadataItem.Create(nameof(this.End), SrcConstants.LockEnd, this.End);
                yield return CellMetadataItem.Create(nameof(this.Format), SrcConstants.LockFormat, this.Format);
                yield return CellMetadataItem.Create(nameof(this.FromGroupFormat), SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
                yield return CellMetadataItem.Create(nameof(this.Group), SrcConstants.LockGroup, this.Group);
                yield return CellMetadataItem.Create(nameof(this.Height), SrcConstants.LockHeight, this.Height);
                yield return CellMetadataItem.Create(nameof(this.MoveX), SrcConstants.LockMoveX, this.MoveX);
                yield return CellMetadataItem.Create(nameof(this.MoveY), SrcConstants.LockMoveY, this.MoveY);
                yield return CellMetadataItem.Create(nameof(this.Rotate), SrcConstants.LockRotate, this.Rotate);
                yield return CellMetadataItem.Create(nameof(this.Select), SrcConstants.LockSelect, this.Select);
                yield return CellMetadataItem.Create(nameof(this.TextEdit), SrcConstants.LockTextEdit, this.TextEdit);
                yield return CellMetadataItem.Create(nameof(this.ThemeColors), SrcConstants.LockThemeColors, this.ThemeColors);
                yield return CellMetadataItem.Create(nameof(this.ThemeEffects), SrcConstants.LockThemeEffects, this.ThemeEffects);
                yield return CellMetadataItem.Create(nameof(this.VertexEdit), SrcConstants.LockVertexEdit, this.VertexEdit);
                yield return CellMetadataItem.Create(nameof(this.Width), SrcConstants.LockWidth, this.Width);
            }
        }


    }
}