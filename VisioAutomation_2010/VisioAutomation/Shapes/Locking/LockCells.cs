using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes.Locking
{
    public class LockCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public ShapeSheet.CellData LockAspect { get; set; }
        public ShapeSheet.CellData LockBegin { get; set; }
        public ShapeSheet.CellData LockCalcWH { get; set; }
        public ShapeSheet.CellData LockCrop { get; set; }
        public ShapeSheet.CellData LockCustProp { get; set; }
        public ShapeSheet.CellData LockDelete { get; set; }
        public ShapeSheet.CellData LockEnd { get; set; }
        public ShapeSheet.CellData LockFormat { get; set; }
        public ShapeSheet.CellData LockFromGroupFormat { get; set; }
        public ShapeSheet.CellData LockGroup { get; set; }
        public ShapeSheet.CellData LockHeight { get; set; }
        public ShapeSheet.CellData LockMoveX { get; set; }
        public ShapeSheet.CellData LockMoveY { get; set; }
        public ShapeSheet.CellData LockRotate { get; set; }
        public ShapeSheet.CellData LockSelect { get; set; }
        public ShapeSheet.CellData LockTextEdit { get; set; }
        public ShapeSheet.CellData LockThemeColors { get; set; }
        public ShapeSheet.CellData LockThemeEffects { get; set; }
        public ShapeSheet.CellData LockVtxEdit { get; set; }
        public ShapeSheet.CellData LockWidth { get; set; }

        public override IEnumerable<SRCFormulaPair> SRCFormulaPairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SrcConstants.LockAspect, this.LockAspect.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockBegin, this.LockBegin.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCalcWH, this.LockCalcWH.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCrop, this.LockCrop.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockCustProp, this.LockCustProp.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockDelete, this.LockDelete.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockEnd, this.LockEnd.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFormat, this.LockFormat.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockFromGroupFormat, this.LockFromGroupFormat.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockGroup, this.LockGroup.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockHeight, this.LockHeight.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveX, this.LockMoveX.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockMoveY, this.LockMoveY.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockRotate, this.LockRotate.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockSelect, this.LockSelect.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockTextEdit, this.LockTextEdit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeColors, this.LockThemeColors.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockThemeEffects, this.LockThemeEffects.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockVtxEdit, this.LockVtxEdit.Formula);
                yield return this.newpair(ShapeSheet.SrcConstants.LockWidth, this.LockWidth.Formula);
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