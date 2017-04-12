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
                yield return this.newpair(ShapeSheet.SrcConstants.LockAspect, this.Aspect.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockBegin, this.Begin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCalcWH, this.CalcWH.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCrop, this.Crop.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCustomProp, this.CustProp.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockDelete, this.Delete.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockEnd, this.End.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFormat, this.Format.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFromGroupFormat, this.FromGroupFormat.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockGroup, this.Group.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockHeight, this.Height.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveX, this.MoveX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveY, this.MoveY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockRotate, this.Rotate.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockSelect, this.Select.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockTextEdit, this.TextEdit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeColors, this.ThemeColors.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeEffects, this.ThemeEffects.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockVertexEdit, this.VertexEdit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockWidth, this.Width.Formula);
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