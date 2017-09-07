using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class LockCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData Aspect { get; set; }
        public ShapeSheet.CellData Begin { get; set; }
        public ShapeSheet.CellData CalcWH { get; set; }
        public ShapeSheet.CellData Crop { get; set; }
        public ShapeSheet.CellData CustProp { get; set; }
        public ShapeSheet.CellData Delete { get; set; }
        public ShapeSheet.CellData End { get; set; }
        public ShapeSheet.CellData Format { get; set; }
        public ShapeSheet.CellData FromGroupFormat { get; set; }
        public ShapeSheet.CellData Group { get; set; }
        public ShapeSheet.CellData Height { get; set; }
        public ShapeSheet.CellData MoveX { get; set; }
        public ShapeSheet.CellData MoveY { get; set; }
        public ShapeSheet.CellData Rotate { get; set; }
        public ShapeSheet.CellData Select { get; set; }
        public ShapeSheet.CellData TextEdit { get; set; }
        public ShapeSheet.CellData ThemeColors { get; set; }
        public ShapeSheet.CellData ThemeEffects { get; set; }
        public ShapeSheet.CellData VertexEdit { get; set; }
        public ShapeSheet.CellData Width { get; set; }

        public override IEnumerable<SrcFormulaPair> SrcFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.LockAspect, this.Aspect.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockBegin, this.Begin.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCalcWH, this.CalcWH.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCrop, this.Crop.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCustomProp, this.CustProp.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockDelete, this.Delete.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockEnd, this.End.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFormat, this.Format.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFromGroupFormat, this.FromGroupFormat.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockGroup, this.Group.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockHeight, this.Height.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveX, this.MoveX.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveY, this.MoveY.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockRotate, this.Rotate.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockSelect, this.Select.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockTextEdit, this.TextEdit.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeColors, this.ThemeColors.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeEffects, this.ThemeEffects.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockVertexEdit, this.VertexEdit.Value);
                yield return this.newpair(ShapeSheet.SrcConstants.LockWidth, this.Width.Value);
            }
        }


        public static List<LockCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids, cvt);
        }

        public static LockCells GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetCellGroup(shape, cvt);
        }

        private static readonly System.Lazy<LockCellsReader> lazy_query = new System.Lazy<LockCellsReader>();
    }
}