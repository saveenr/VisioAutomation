using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries

{
    class LockCellsQuery : CellGroupSingleRowQuery<Shapes.LockCells, double>
    {
        public ColumnQuery LockAspect { get; set; }
        public ColumnQuery LockBegin { get; set; }
        public ColumnQuery LockCalcWH { get; set; }
        public ColumnQuery LockCrop { get; set; }
        public ColumnQuery LockCustProp { get; set; }
        public ColumnQuery LockDelete { get; set; }
        public ColumnQuery LockEnd { get; set; }
        public ColumnQuery LockFormat { get; set; }
        public ColumnQuery LockFromGroupFormat { get; set; }
        public ColumnQuery LockGroup { get; set; }
        public ColumnQuery LockHeight { get; set; }
        public ColumnQuery LockMoveX { get; set; }
        public ColumnQuery LockMoveY { get; set; }
        public ColumnQuery LockRotate { get; set; }
        public ColumnQuery LockSelect { get; set; }
        public ColumnQuery LockTextEdit { get; set; }
        public ColumnQuery LockThemeColors { get; set; }
        public ColumnQuery LockThemeEffects { get; set; }
        public ColumnQuery LockVtxEdit { get; set; }
        public ColumnQuery LockWidth { get; set; }

        public LockCellsQuery()
        {





            this.LockAspect = this.query.AddCell(SRCCON.LockAspect, nameof(SRCCON.LockAspect));
            this.LockBegin = this.query.AddCell(SRCCON.LockBegin, nameof(SRCCON.LockBegin));
            this.LockCalcWH = this.query.AddCell(SRCCON.LockCalcWH, nameof(SRCCON.LockCalcWH));
            this.LockCrop = this.query.AddCell(SRCCON.LockCrop, nameof(SRCCON.LockCrop));
            this.LockCustProp = this.query.AddCell(SRCCON.LockCustProp, nameof(SRCCON.LockCustProp));
            this.LockDelete = this.query.AddCell(SRCCON.LockDelete, nameof(SRCCON.LockDelete));
            this.LockEnd = this.query.AddCell(SRCCON.LockEnd, nameof(SRCCON.LockEnd));
            this.LockFormat = this.query.AddCell(SRCCON.LockFormat, nameof(SRCCON.LockFormat));
            this.LockFromGroupFormat = this.query.AddCell(SRCCON.LockFromGroupFormat, nameof(SRCCON.LockFromGroupFormat));
            this.LockGroup = this.query.AddCell(SRCCON.LockGroup, nameof(SRCCON.LockGroup));
            this.LockHeight = this.query.AddCell(SRCCON.LockHeight, nameof(SRCCON.LockHeight));
            this.LockMoveX = this.query.AddCell(SRCCON.LockMoveX, nameof(SRCCON.LockMoveX));
            this.LockMoveY = this.query.AddCell(SRCCON.LockMoveY, nameof(SRCCON.LockMoveY));
            this.LockRotate = this.query.AddCell(SRCCON.LockRotate, nameof(SRCCON.LockRotate));
            this.LockSelect = this.query.AddCell(SRCCON.LockSelect, nameof(SRCCON.LockSelect));
            this.LockTextEdit = this.query.AddCell(SRCCON.LockTextEdit, nameof(SRCCON.LockTextEdit));
            this.LockThemeColors = this.query.AddCell(SRCCON.LockThemeColors, nameof(SRCCON.LockThemeColors));
            this.LockThemeEffects = this.query.AddCell(SRCCON.LockThemeEffects, nameof(SRCCON.LockThemeEffects));
            this.LockVtxEdit = this.query.AddCell(SRCCON.LockVtxEdit, nameof(SRCCON.LockVtxEdit));
            this.LockWidth = this.query.AddCell(SRCCON.LockWidth, nameof(SRCCON.LockWidth));


        }

        public override Shapes.LockCells CellDataToCellGroup(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Shapes.LockCells();
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