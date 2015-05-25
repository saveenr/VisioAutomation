using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using VAQUERY=VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class LockCells : ShapeSheet.CellGroups.CellGroup
    {
        public ShapeSheet.CellData<bool> LockAspect { get; set; }
        public ShapeSheet.CellData<bool> LockBegin { get; set; }
        public ShapeSheet.CellData<bool> LockCalcWH { get; set; }
        public ShapeSheet.CellData<bool> LockCrop { get; set; }
        public ShapeSheet.CellData<bool> LockCustProp { get; set; }
        public ShapeSheet.CellData<bool> LockDelete { get; set; }
        public ShapeSheet.CellData<bool> LockEnd { get; set; }
        public ShapeSheet.CellData<bool> LockFormat { get; set; }
        public ShapeSheet.CellData<bool> LockFromGroupFormat { get; set; }
        public ShapeSheet.CellData<bool> LockGroup { get; set; }
        public ShapeSheet.CellData<bool> LockHeight { get; set; }
        public ShapeSheet.CellData<bool> LockMoveX { get; set; }
        public ShapeSheet.CellData<bool> LockMoveY { get; set; }
        public ShapeSheet.CellData<bool> LockRotate { get; set; }
        public ShapeSheet.CellData<bool> LockSelect { get; set; }
        public ShapeSheet.CellData<bool> LockTextEdit { get; set; }
        public ShapeSheet.CellData<bool> LockThemeColors { get; set; }
        public ShapeSheet.CellData<bool> LockThemeEffects { get; set; }
        public ShapeSheet.CellData<bool> LockVtxEdit { get; set; }
        public ShapeSheet.CellData<bool> LockWidth { get; set; }

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
            return ShapeSheet.CellGroups.CellGroup._GetCells<LockCells, double>(page, shapeids, query, query.GetCells);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = LockCells.lazy_query.Value;
            return ShapeSheet.CellGroups.CellGroup._GetCells<LockCells, double>(shape, query, query.GetCells);
        }

        private static System.Lazy<ShapeSheet.Query.Common.LockCellsQuery> lazy_query = new System.Lazy<ShapeSheet.Query.Common.LockCellsQuery>();


    }
}