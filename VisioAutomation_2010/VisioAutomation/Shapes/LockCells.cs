using System.Collections.Generic;
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
                yield return this.newpair(ShapeSheet.SrcConstants.LockAspect, this.Aspect.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockBegin, this.Begin.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCalcWH, this.CalcWH.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCrop, this.Crop.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCustomProp, this.CustProp.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockDelete, this.Delete.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockEnd, this.End.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFormat, this.Format.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFromGroupFormat, this.FromGroupFormat.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockGroup, this.Group.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockHeight, this.Height.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveX, this.MoveX.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveY, this.MoveY.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockRotate, this.Rotate.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockSelect, this.Select.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockTextEdit, this.TextEdit.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeColors, this.ThemeColors.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeEffects, this.ThemeEffects.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockVertexEdit, this.VertexEdit.ValueF);
                yield return this.newpair(ShapeSheet.SrcConstants.LockWidth, this.Width.ValueF);
            }
        }


        public static List<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static readonly System.Lazy<LockCellsReader> lazy_query = new System.Lazy<LockCellsReader>();
    }
}