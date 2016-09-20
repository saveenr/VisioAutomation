using IVisio = Microsoft.Office.Interop.Visio;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.CellGroups.Queries;

namespace VisioAutomation.Shapes
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

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return this.newpair(ShapeSheet.SRCConstants.LockAspect, this.LockAspect.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockBegin, this.LockBegin.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockCalcWH, this.LockCalcWH.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockCrop, this.LockCrop.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockCustProp, this.LockCustProp.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockDelete, this.LockDelete.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockEnd, this.LockEnd.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockFormat, this.LockFormat.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockFromGroupFormat, this.LockFromGroupFormat.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockGroup, this.LockGroup.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockHeight, this.LockHeight.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockMoveX, this.LockMoveX.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockMoveY, this.LockMoveY.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockRotate, this.LockRotate.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockSelect, this.LockSelect.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockTextEdit, this.LockTextEdit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockThemeColors, this.LockThemeColors.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockThemeEffects, this.LockThemeEffects.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockVtxEdit, this.LockVtxEdit.Formula);
                yield return this.newpair(ShapeSheet.SRCConstants.LockWidth, this.LockWidth.Formula);
            }
        }


        public static IList<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetCellGroups(page, shapeids);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = LockCells.lazy_query.Value;
            return query.GetCellGroup(shape);
        }

        private static System.Lazy<LockCellsQuery> lazy_query = new System.Lazy<LockCellsQuery>();


    }
}