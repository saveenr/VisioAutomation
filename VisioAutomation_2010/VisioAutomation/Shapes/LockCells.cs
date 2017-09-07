using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class LockCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral Aspect { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Begin { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CalcWH { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Crop { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral CustProp { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Delete { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral End { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Format { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FromGroupFormat { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Group { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Height { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral MoveX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral MoveY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Rotate { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Select { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral TextEdit { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ThemeColors { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ThemeEffects { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral VertexEdit { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral Width { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockAspect, this.Aspect.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockBegin, this.Begin.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockCalcWH, this.CalcWH.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockCrop, this.Crop.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockCustomProp, this.CustProp.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockDelete, this.Delete.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockEnd, this.End.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockFormat, this.Format.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockFromGroupFormat, this.FromGroupFormat.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockGroup, this.Group.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockHeight, this.Height.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockMoveX, this.MoveX.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockMoveY, this.MoveY.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockRotate, this.Rotate.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockSelect, this.Select.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockTextEdit, this.TextEdit.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockThemeColors, this.ThemeColors.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockThemeEffects, this.ThemeEffects.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockVertexEdit, this.VertexEdit.Value);
                yield return SrcFormulaPair.Create(ShapeSheet.SrcConstants.LockWidth, this.Width.Value);
            }
        }


        public static List<LockCells> GetFormulas(IVisio.Page page, IList<int> shapeids)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetFormulas(page, shapeids);
        }

        public static List<LockCells> GetResults(IVisio.Page page, IList<int> shapeids)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetResults(page, shapeids);
        }

        public static LockCells GetFormulas(IVisio.Shape shape)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetFormulas(shape);
        }

        public static LockCells GetResults(IVisio.Shape shape, CellValueType cvt)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetResults(shape);
        }

        private static readonly System.Lazy<LockCellsReader> lazy_query = new System.Lazy<LockCellsReader>();
    }
}