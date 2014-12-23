using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;

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

        public override IEnumerable<VA.ShapeSheet.CellGroups.BaseCellGroup.SRCValuePair> EnumPairs()
        {
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockAspect, this.LockAspect.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockBegin, this.LockBegin.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockCalcWH, this.LockCalcWH.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockCrop, this.LockCrop.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockCustProp, this.LockCustProp.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockDelete, this.LockDelete.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockEnd, this.LockEnd.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockFormat, this.LockFormat.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockFromGroupFormat, this.LockFromGroupFormat.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockGroup, this.LockGroup.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockHeight, this.LockHeight.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockMoveX, this.LockMoveX.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockMoveY, this.LockMoveY.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockRotate, this.LockRotate.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockSelect, this.LockSelect.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockTextEdit, this.LockTextEdit.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockThemeColors, this.LockThemeColors.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockThemeEffects, this.LockThemeEffects.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockVtxEdit, this.LockVtxEdit.Formula);
            yield return srcvaluepair(ShapeSheet.SRCConstants.LockWidth, this.LockWidth.Formula);
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
                this.LockAspect = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockAspect, "LockAspect");
                this.LockBegin = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockBegin, "LockBegin");
                this.LockCalcWH = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockCalcWH, "LockCalcWH");
                this.LockCrop = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockCrop, "LockCrop");
                this.LockCustProp = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockCustProp, "LockCustProp");
                this.LockDelete = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockDelete, "LockDelete");
                this.LockEnd = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockEnd, "LockEnd");
                this.LockFormat = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockFormat, "LockFormat");
                this.LockFromGroupFormat = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockFromGroupFormat, "LockFromGroupFormat");
                this.LockGroup = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockGroup, "LockGroup");
                this.LockHeight = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockHeight, "LockHeight");
                this.LockMoveX = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockMoveX, "LockMoveX");
                this.LockMoveY = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockMoveY, "LockMoveY");
                this.LockRotate = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockRotate, "LockRotate");
                this.LockSelect = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockSelect, "LockSelect");
                this.LockTextEdit = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockTextEdit, "LockTextEdit");
                this.LockThemeColors = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockThemeColors, "LockThemeColors");
                this.LockThemeEffects = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockThemeEffects, "LockThemeEffects");
                this.LockVtxEdit = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockVtxEdit, "LockVtxEdit");
                this.LockWidth = this.Columns.Add(VA.ShapeSheet.SRCConstants.LockWidth, "LockWidth");
            }

            public LockCells GetCells(VA.ShapeSheet.CellData<double>[] row)
            {
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