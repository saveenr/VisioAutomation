using System.Linq;
using VisioAutomation.Extensions;
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

        private static LockCells get_cells_from_row(LockQuery query, VA.ShapeSheet.Data.Table<VA.ShapeSheet.CellData<double>> table, int row)
        {
            var cells = new LockCells();
            cells.LockAspect = table[row,query.LockAspect].ToBool();
            cells.LockBegin = table[row,query.LockBegin].ToBool();
            cells.LockCalcWH = table[row,query.LockCalcWH].ToBool();
            cells.LockCrop = table[row,query.LockCrop].ToBool();
            cells.LockCustProp = table[row,query.LockCustProp].ToBool();
            cells.LockDelete = table[row,query.LockDelete].ToBool();
            cells.LockEnd = table[row,query.LockEnd].ToBool();
            cells.LockFormat = table[row,query.LockFormat].ToBool();
            cells.LockFromGroupFormat = table[row,query.LockFromGroupFormat].ToBool();
            cells.LockGroup = table[row,query.LockGroup].ToBool();
            cells.LockHeight = table[row,query.LockHeight].ToBool();
            cells.LockMoveX = table[row,query.LockMoveX].ToBool();
            cells.LockMoveY = table[row,query.LockMoveY].ToBool();
            cells.LockRotate = table[row,query.LockRotate].ToBool();
            cells.LockSelect = table[row,query.LockSelect].ToBool();
            cells.LockTextEdit = table[row,query.LockTextEdit].ToBool();
            cells.LockThemeColors = table[row,query.LockThemeColors].ToBool();
            cells.LockThemeEffects = table[row,query.LockThemeEffects].ToBool();
            cells.LockVtxEdit = table[row,query.LockVtxEdit].ToBool();
            cells.LockWidth = table[row,query.LockWidth].ToBool();
            return cells;
        }

        public static IList<LockCells> GetCells(IVisio.Page page, IList<int> shapeids)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRows(page, shapeids, query, get_cells_from_row);
        }

        public static LockCells GetCells(IVisio.Shape shape)
        {
            var query = get_query();
            return VA.ShapeSheet.CellGroups.CellGroup.CellsFromRow(shape, query, get_cells_from_row);
        }

        private static LockQuery m_query;
        private static LockQuery get_query()
        {
            m_query = m_query ?? new LockQuery();
            return m_query;
        }

        class LockQuery : VA.ShapeSheet.Query.CellQuery
        {
            public VA.ShapeSheet.Query.QueryColumn LockAspect { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockBegin { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockCalcWH { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockCrop { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockCustProp { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockDelete { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockEnd { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockFormat { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockFromGroupFormat { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockGroup { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockHeight { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockMoveX { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockMoveY { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockRotate { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockSelect { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockTextEdit { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockThemeColors { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockThemeEffects { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockVtxEdit { get; set; }
            public VA.ShapeSheet.Query.QueryColumn LockWidth { get; set; }

            public LockQuery() :
                base()
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
        }
    }
}