using System.Linq;
using VisioAutomation.Extensions;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

namespace VisioAutomation.Layout
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

        public override void ApplyFormulas(ApplyFormula func)
        {
            func(ShapeSheet.SRCConstants.LockAspect, this.LockAspect.Formula);
            func(ShapeSheet.SRCConstants.LockBegin, this.LockBegin.Formula);
            func(ShapeSheet.SRCConstants.LockCalcWH, this.LockCalcWH.Formula);
            func(ShapeSheet.SRCConstants.LockCrop, this.LockCrop.Formula);
            func(ShapeSheet.SRCConstants.LockCustProp, this.LockCustProp.Formula);
            func(ShapeSheet.SRCConstants.LockDelete, this.LockDelete.Formula);
            func(ShapeSheet.SRCConstants.LockEnd, this.LockEnd.Formula);
            func(ShapeSheet.SRCConstants.LockFormat, this.LockFormat.Formula);
            func(ShapeSheet.SRCConstants.LockFromGroupFormat, this.LockFromGroupFormat.Formula);
            func(ShapeSheet.SRCConstants.LockGroup, this.LockGroup.Formula);
            func(ShapeSheet.SRCConstants.LockHeight, this.LockHeight.Formula);
            func(ShapeSheet.SRCConstants.LockMoveX, this.LockMoveX.Formula);
            func(ShapeSheet.SRCConstants.LockMoveY, this.LockMoveY.Formula);
            func(ShapeSheet.SRCConstants.LockRotate, this.LockRotate.Formula);
            func(ShapeSheet.SRCConstants.LockSelect, this.LockSelect.Formula);
            func(ShapeSheet.SRCConstants.LockTextEdit, this.LockTextEdit.Formula);
            func(ShapeSheet.SRCConstants.LockThemeColors, this.LockThemeColors.Formula);
            func(ShapeSheet.SRCConstants.LockThemeEffects, this.LockThemeEffects.Formula);
            func(ShapeSheet.SRCConstants.LockVtxEdit, this.LockVtxEdit.Formula);
            func(ShapeSheet.SRCConstants.LockWidth, this.LockWidth.Formula);
        }


        public static IList<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(page, shapeids, query, query.GetCells);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup._GetCells(shape, query, query.GetCells);
        }


        private static LockCellQuery _mCellQuery;
        private static LockCellQuery get_query()
        {
            _mCellQuery = _mCellQuery ?? new LockCellQuery();
            return _mCellQuery;
        }

        class LockCellQuery : VA.ShapeSheet.Query.CellQuery
        {
            public Column LockAspect { get; set; }
            public Column LockBegin { get; set; }
            public Column LockCalcWH { get; set; }
            public Column LockCrop { get; set; }
            public Column LockCustProp { get; set; }
            public Column LockDelete { get; set; }
            public Column LockEnd { get; set; }
            public Column LockFormat { get; set; }
            public Column LockFromGroupFormat { get; set; }
            public Column LockGroup { get; set; }
            public Column LockHeight { get; set; }
            public Column LockMoveX { get; set; }
            public Column LockMoveY { get; set; }
            public Column LockRotate { get; set; }
            public Column LockSelect { get; set; }
            public Column LockTextEdit { get; set; }
            public Column LockThemeColors { get; set; }
            public Column LockThemeEffects { get; set; }
            public Column LockVtxEdit { get; set; }
            public Column LockWidth { get; set; }

            public LockCellQuery() 
            {
                this.LockAspect = this.AddColumn(VA.ShapeSheet.SRCConstants.LockAspect, "LockAspect");
                this.LockBegin = this.AddColumn(VA.ShapeSheet.SRCConstants.LockBegin, "LockBegin");
                this.LockCalcWH = this.AddColumn(VA.ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH");
                this.LockCrop = this.AddColumn(VA.ShapeSheet.SRCConstants.LockCrop, "LockCrop");
                this.LockCustProp = this.AddColumn(VA.ShapeSheet.SRCConstants.LockCustProp, "LockCustProp");
                this.LockDelete = this.AddColumn(VA.ShapeSheet.SRCConstants.LockDelete, "LockDelete");
                this.LockEnd = this.AddColumn(VA.ShapeSheet.SRCConstants.LockEnd, "LockEnd");
                this.LockFormat = this.AddColumn(VA.ShapeSheet.SRCConstants.LockFormat, "LockFormat");
                this.LockFromGroupFormat = this.AddColumn(VA.ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat");
                this.LockGroup = this.AddColumn(VA.ShapeSheet.SRCConstants.LockGroup, "LockGroup");
                this.LockHeight = this.AddColumn(VA.ShapeSheet.SRCConstants.LockHeight, "LockHeight");
                this.LockMoveX = this.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveX, "LockMoveX");
                this.LockMoveY = this.AddColumn(VA.ShapeSheet.SRCConstants.LockMoveY, "LockMoveY");
                this.LockRotate = this.AddColumn(VA.ShapeSheet.SRCConstants.LockRotate, "LockRotate");
                this.LockSelect = this.AddColumn(VA.ShapeSheet.SRCConstants.LockSelect, "LockSelect");
                this.LockTextEdit = this.AddColumn(VA.ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit");
                this.LockThemeColors = this.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors");
                this.LockThemeEffects = this.AddColumn(VA.ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects");
                this.LockVtxEdit = this.AddColumn(VA.ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit");
                this.LockWidth = this.AddColumn(VA.ShapeSheet.SRCConstants.LockWidth, "LockWidth");
            }

            public LockCells GetCells(QueryResult<CellData<double>> data_for_shape)
            {
                var row = data_for_shape.Cells;
                var cells = new LockCells();
                cells.LockAspect = row[this.LockAspect.Ordinal].ToBool();
                cells.LockBegin = row[this.LockBegin.Ordinal].ToBool();
                cells.LockCalcWH = row[this.LockCalcWH.Ordinal].ToBool();
                cells.LockCrop = row[this.LockCrop.Ordinal].ToBool();
                cells.LockCustProp = row[this.LockCustProp.Ordinal].ToBool();
                cells.LockDelete = row[this.LockDelete.Ordinal].ToBool();
                cells.LockEnd = row[this.LockEnd.Ordinal].ToBool();
                cells.LockFormat = row[this.LockFormat.Ordinal].ToBool();
                cells.LockFromGroupFormat = row[this.LockFromGroupFormat.Ordinal].ToBool();
                cells.LockGroup = row[this.LockGroup.Ordinal].ToBool();
                cells.LockHeight = row[this.LockHeight.Ordinal].ToBool();
                cells.LockMoveX = row[this.LockMoveX.Ordinal].ToBool();
                cells.LockMoveY = row[this.LockMoveY.Ordinal].ToBool();
                cells.LockRotate = row[this.LockRotate.Ordinal].ToBool();
                cells.LockSelect = row[this.LockSelect.Ordinal].ToBool();
                cells.LockTextEdit = row[this.LockTextEdit.Ordinal].ToBool();
                cells.LockThemeColors = row[this.LockThemeColors.Ordinal].ToBool();
                cells.LockThemeEffects = row[this.LockThemeEffects.Ordinal].ToBool();
                cells.LockVtxEdit = row[this.LockVtxEdit.Ordinal].ToBool();
                cells.LockWidth = row[this.LockWidth.Ordinal].ToBool();
                return cells;
            }
        }
    }
}