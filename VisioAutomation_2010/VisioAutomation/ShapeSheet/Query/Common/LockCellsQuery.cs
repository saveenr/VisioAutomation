using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class LockCellsQuery : CellQuery
    {
        public Query.CellColumn LockAspect { get; set; }
        public Query.CellColumn LockBegin { get; set; }
        public Query.CellColumn LockCalcWH { get; set; }
        public Query.CellColumn LockCrop { get; set; }
        public Query.CellColumn LockCustProp { get; set; }
        public Query.CellColumn LockDelete { get; set; }
        public Query.CellColumn LockEnd { get; set; }
        public Query.CellColumn LockFormat { get; set; }
        public Query.CellColumn LockFromGroupFormat { get; set; }
        public Query.CellColumn LockGroup { get; set; }
        public Query.CellColumn LockHeight { get; set; }
        public Query.CellColumn LockMoveX { get; set; }
        public Query.CellColumn LockMoveY { get; set; }
        public Query.CellColumn LockRotate { get; set; }
        public Query.CellColumn LockSelect { get; set; }
        public Query.CellColumn LockTextEdit { get; set; }
        public Query.CellColumn LockThemeColors { get; set; }
        public Query.CellColumn LockThemeEffects { get; set; }
        public Query.CellColumn LockVtxEdit { get; set; }
        public Query.CellColumn LockWidth { get; set; }

        public LockCellsQuery()
        {





            this.LockAspect = this.AddCell(SRCCON.LockAspect, nameof(SRCCON.LockAspect));
            this.LockBegin = this.AddCell(SRCCON.LockBegin, nameof(SRCCON.LockBegin));
            this.LockCalcWH = this.AddCell(SRCCON.LockCalcWH, nameof(SRCCON.LockCalcWH));
            this.LockCrop = this.AddCell(SRCCON.LockCrop, nameof(SRCCON.LockCrop));
            this.LockCustProp = this.AddCell(SRCCON.LockCustProp, nameof(SRCCON.LockCustProp));
            this.LockDelete = this.AddCell(SRCCON.LockDelete, nameof(SRCCON.LockDelete));
            this.LockEnd = this.AddCell(SRCCON.LockEnd, nameof(SRCCON.LockEnd));
            this.LockFormat = this.AddCell(SRCCON.LockFormat, nameof(SRCCON.LockFormat));
            this.LockFromGroupFormat = this.AddCell(SRCCON.LockFromGroupFormat, nameof(SRCCON.LockFromGroupFormat));
            this.LockGroup = this.AddCell(SRCCON.LockGroup, nameof(SRCCON.LockGroup));
            this.LockHeight = this.AddCell(SRCCON.LockHeight, nameof(SRCCON.LockHeight));
            this.LockMoveX = this.AddCell(SRCCON.LockMoveX, nameof(SRCCON.LockMoveX));
            this.LockMoveY = this.AddCell(SRCCON.LockMoveY, nameof(SRCCON.LockMoveY));
            this.LockRotate = this.AddCell(SRCCON.LockRotate, nameof(SRCCON.LockRotate));
            this.LockSelect = this.AddCell(SRCCON.LockSelect, nameof(SRCCON.LockSelect));
            this.LockTextEdit = this.AddCell(SRCCON.LockTextEdit, nameof(SRCCON.LockTextEdit));
            this.LockThemeColors = this.AddCell(SRCCON.LockThemeColors, nameof(SRCCON.LockThemeColors));
            this.LockThemeEffects = this.AddCell(SRCCON.LockThemeEffects, nameof(SRCCON.LockThemeEffects));
            this.LockVtxEdit = this.AddCell(SRCCON.LockVtxEdit, nameof(SRCCON.LockVtxEdit));
            this.LockWidth = this.AddCell(SRCCON.LockWidth, nameof(SRCCON.LockWidth));


        }

        public VisioAutomation.Shapes.LockCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.LockCells();
            cells.LockAspect = Extensions.CellDataMethods.ToBool(row[this.LockAspect]);
            cells.LockBegin = Extensions.CellDataMethods.ToBool(row[this.LockBegin]);
            cells.LockCalcWH = Extensions.CellDataMethods.ToBool(row[this.LockCalcWH]);
            cells.LockCrop = Extensions.CellDataMethods.ToBool(row[this.LockCrop]);
            cells.LockCustProp = Extensions.CellDataMethods.ToBool(row[this.LockCustProp]);
            cells.LockDelete = Extensions.CellDataMethods.ToBool(row[this.LockDelete]);
            cells.LockEnd = Extensions.CellDataMethods.ToBool(row[this.LockEnd]);
            cells.LockFormat = Extensions.CellDataMethods.ToBool(row[this.LockFormat]);
            cells.LockFromGroupFormat = Extensions.CellDataMethods.ToBool(row[this.LockFromGroupFormat]);
            cells.LockGroup = Extensions.CellDataMethods.ToBool(row[this.LockGroup]);
            cells.LockHeight = Extensions.CellDataMethods.ToBool(row[this.LockHeight]);
            cells.LockMoveX = Extensions.CellDataMethods.ToBool(row[this.LockMoveX]);
            cells.LockMoveY = Extensions.CellDataMethods.ToBool(row[this.LockMoveY]);
            cells.LockRotate = Extensions.CellDataMethods.ToBool(row[this.LockRotate]);
            cells.LockSelect = Extensions.CellDataMethods.ToBool(row[this.LockSelect]);
            cells.LockTextEdit = Extensions.CellDataMethods.ToBool(row[this.LockTextEdit]);
            cells.LockThemeColors = Extensions.CellDataMethods.ToBool(row[this.LockThemeColors]);
            cells.LockThemeEffects = Extensions.CellDataMethods.ToBool(row[this.LockThemeEffects]);
            cells.LockVtxEdit = Extensions.CellDataMethods.ToBool(row[this.LockVtxEdit]);
            cells.LockWidth = Extensions.CellDataMethods.ToBool(row[this.LockWidth]);
            return cells;
        }
    }
}