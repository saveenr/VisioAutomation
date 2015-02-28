using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class LockCells : VA.ShapeSheet.CellGroups.CellGroup
    {
        public VA.ShapeSheet.CellData<bool> LockAspect { get; set; }
        public VA.ShapeSheet.CellData<bool> LockBegin { get; set; }
        public VA.ShapeSheet.CellData<bool> LockCalcWH { get; set; }
        public VA.ShapeSheet.CellData<bool> LockCrop { get; set; }
        public VA.ShapeSheet.CellData<bool> LockCustProp { get; set; }
        public VA.ShapeSheet.CellData<bool> LockDelete { get; set; }
        public VA.ShapeSheet.CellData<bool> LockEnd { get; set; }
        public VA.ShapeSheet.CellData<bool> LockFormat { get; set; }
        public VA.ShapeSheet.CellData<bool> LockFromGroupFormat { get; set; }
        public VA.ShapeSheet.CellData<bool> LockGroup { get; set; }
        public VA.ShapeSheet.CellData<bool> LockHeight { get; set; }
        public VA.ShapeSheet.CellData<bool> LockMoveX { get; set; }
        public VA.ShapeSheet.CellData<bool> LockMoveY { get; set; }
        public VA.ShapeSheet.CellData<bool> LockRotate { get; set; }
        public VA.ShapeSheet.CellData<bool> LockSelect { get; set; }
        public VA.ShapeSheet.CellData<bool> LockTextEdit { get; set; }
        public VA.ShapeSheet.CellData<bool> LockThemeColors { get; set; }
        public VA.ShapeSheet.CellData<bool> LockThemeEffects { get; set; }
        public VA.ShapeSheet.CellData<bool> LockVtxEdit { get; set; }
        public VA.ShapeSheet.CellData<bool> LockWidth { get; set; }

        public override IEnumerable<SRCFormulaPair> Pairs
        {
            get
            {
                yield return newpair(ShapeSheet.SRCConstants.LockAspect, this.LockAspect.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockBegin, this.LockBegin.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockCalcWH, this.LockCalcWH.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockCrop, this.LockCrop.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockCustProp, this.LockCustProp.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockDelete, this.LockDelete.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockEnd, this.LockEnd.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockFormat, this.LockFormat.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockFromGroupFormat, this.LockFromGroupFormat.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockGroup, this.LockGroup.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockHeight, this.LockHeight.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockMoveX, this.LockMoveX.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockMoveY, this.LockMoveY.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockRotate, this.LockRotate.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockSelect, this.LockSelect.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockTextEdit, this.LockTextEdit.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockThemeColors, this.LockThemeColors.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockThemeEffects, this.LockThemeEffects.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockVtxEdit, this.LockVtxEdit.Formula);
                yield return newpair(ShapeSheet.SRCConstants.LockWidth, this.LockWidth.Formula);
            }
        }


        public static IList<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<LockCells, double>(page, shapeids, query, query.GetCells);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells<LockCells, double>(shape, query, query.GetCells);
        }


        private static LockCellQuery _mCellQuery;
        private static LockCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new LockCellQuery();
            return _mCellQuery;
        }

        class LockCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public CellColumn LockAspect { get; set; }
            public CellColumn LockBegin { get; set; }
            public CellColumn LockCalcWH { get; set; }
            public CellColumn LockCrop { get; set; }
            public CellColumn LockCustProp { get; set; }
            public CellColumn LockDelete { get; set; }
            public CellColumn LockEnd { get; set; }
            public CellColumn LockFormat { get; set; }
            public CellColumn LockFromGroupFormat { get; set; }
            public CellColumn LockGroup { get; set; }
            public CellColumn LockHeight { get; set; }
            public CellColumn LockMoveX { get; set; }
            public CellColumn LockMoveY { get; set; }
            public CellColumn LockRotate { get; set; }
            public CellColumn LockSelect { get; set; }
            public CellColumn LockTextEdit { get; set; }
            public CellColumn LockThemeColors { get; set; }
            public CellColumn LockThemeEffects { get; set; }
            public CellColumn LockVtxEdit { get; set; }
            public CellColumn LockWidth { get; set; }

            public LockCellQuery() 
            {
                this.LockAspect = this.AddCell(VA.ShapeSheet.SRCConstants.LockAspect);
                this.LockBegin = this.AddCell(VA.ShapeSheet.SRCConstants.LockBegin);
                this.LockCalcWH = this.AddCell(VA.ShapeSheet.SRCConstants.LockCalcWH);
                this.LockCrop = this.AddCell(VA.ShapeSheet.SRCConstants.LockCrop);
                this.LockCustProp = this.AddCell(VA.ShapeSheet.SRCConstants.LockCustProp);
                this.LockDelete = this.AddCell(VA.ShapeSheet.SRCConstants.LockDelete);
                this.LockEnd = this.AddCell(VA.ShapeSheet.SRCConstants.LockEnd);
                this.LockFormat = this.AddCell(VA.ShapeSheet.SRCConstants.LockFormat);
                this.LockFromGroupFormat = this.AddCell(VA.ShapeSheet.SRCConstants.LockFromGroupFormat);
                this.LockGroup = this.AddCell(VA.ShapeSheet.SRCConstants.LockGroup);
                this.LockHeight = this.AddCell(VA.ShapeSheet.SRCConstants.LockHeight);
                this.LockMoveX = this.AddCell(VA.ShapeSheet.SRCConstants.LockMoveX);
                this.LockMoveY = this.AddCell(VA.ShapeSheet.SRCConstants.LockMoveY);
                this.LockRotate = this.AddCell(VA.ShapeSheet.SRCConstants.LockRotate);
                this.LockSelect = this.AddCell(VA.ShapeSheet.SRCConstants.LockSelect);
                this.LockTextEdit = this.AddCell(VA.ShapeSheet.SRCConstants.LockTextEdit);
                this.LockThemeColors = this.AddCell(VA.ShapeSheet.SRCConstants.LockThemeColors);
                this.LockThemeEffects = this.AddCell(VA.ShapeSheet.SRCConstants.LockThemeEffects);
                this.LockVtxEdit = this.AddCell(VA.ShapeSheet.SRCConstants.LockVtxEdit);
                this.LockWidth = this.AddCell(VA.ShapeSheet.SRCConstants.LockWidth);
            }

            public LockCells GetCells(IList<VA.ShapeSheet.CellData<double>> row)
            {
                var cells = new LockCells();
                cells.LockAspect = row[this.LockAspect].ToBool();
                cells.LockBegin = row[this.LockBegin].ToBool();
                cells.LockCalcWH = row[this.LockCalcWH].ToBool();
                cells.LockCrop = row[this.LockCrop].ToBool();
                cells.LockCustProp = row[this.LockCustProp].ToBool();
                cells.LockDelete = row[this.LockDelete].ToBool();
                cells.LockEnd = row[this.LockEnd].ToBool();
                cells.LockFormat = row[this.LockFormat].ToBool();
                cells.LockFromGroupFormat = row[this.LockFromGroupFormat].ToBool();
                cells.LockGroup = row[this.LockGroup].ToBool();
                cells.LockHeight = row[this.LockHeight].ToBool();
                cells.LockMoveX = row[this.LockMoveX].ToBool();
                cells.LockMoveY = row[this.LockMoveY].ToBool();
                cells.LockRotate = row[this.LockRotate].ToBool();
                cells.LockSelect = row[this.LockSelect].ToBool();
                cells.LockTextEdit = row[this.LockTextEdit].ToBool();
                cells.LockThemeColors = row[this.LockThemeColors].ToBool();
                cells.LockThemeEffects = row[this.LockThemeEffects].ToBool();
                cells.LockVtxEdit = row[this.LockVtxEdit].ToBool();
                cells.LockWidth = row[this.LockWidth].ToBool();
                return cells;
            }
        }
    }
}