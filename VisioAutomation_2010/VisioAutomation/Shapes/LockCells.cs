using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;
using VisioAutomation.ShapeSheet.Query;

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
            var query = get_query();
            return _GetCells<LockCells, double>(page, shapeids, query, query.GetCells);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return _GetCells<LockCells, double>(shape, query, query.GetCells);
        }


        private static LockCellQuery _mCellQuery;
        private static LockCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new LockCellQuery();
            return _mCellQuery;
        }

        class LockCellQuery : CellQuery
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
                this.LockAspect = this.AddCell(ShapeSheet.SRCConstants.LockAspect, "LockAspect");
                this.LockBegin = this.AddCell(ShapeSheet.SRCConstants.LockBegin, "LockBegin");
                this.LockCalcWH = this.AddCell(ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH");
                this.LockCrop = this.AddCell(ShapeSheet.SRCConstants.LockCrop, "LockCrop");
                this.LockCustProp = this.AddCell(ShapeSheet.SRCConstants.LockCustProp, "LockCustProp");
                this.LockDelete = this.AddCell(ShapeSheet.SRCConstants.LockDelete, "LockDelete");
                this.LockEnd = this.AddCell(ShapeSheet.SRCConstants.LockEnd, "LockEnd");
                this.LockFormat = this.AddCell(ShapeSheet.SRCConstants.LockFormat, "LockFormat");
                this.LockFromGroupFormat = this.AddCell(ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat");
                this.LockGroup = this.AddCell(ShapeSheet.SRCConstants.LockGroup, "LockGroup");
                this.LockHeight = this.AddCell(ShapeSheet.SRCConstants.LockHeight, "LockHeight");
                this.LockMoveX = this.AddCell(ShapeSheet.SRCConstants.LockMoveX, "LockMoveX");
                this.LockMoveY = this.AddCell(ShapeSheet.SRCConstants.LockMoveY, "LockMoveY");
                this.LockRotate = this.AddCell(ShapeSheet.SRCConstants.LockRotate, "LockRotate");
                this.LockSelect = this.AddCell(ShapeSheet.SRCConstants.LockSelect, "LockSelect");
                this.LockTextEdit = this.AddCell(ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit");
                this.LockThemeColors = this.AddCell(ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors");
                this.LockThemeEffects = this.AddCell(ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects");
                this.LockVtxEdit = this.AddCell(ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit");
                this.LockWidth = this.AddCell(ShapeSheet.SRCConstants.LockWidth, "LockWidth");

            }

            public LockCells GetCells(IList<ShapeSheet.CellData<double>> row)
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