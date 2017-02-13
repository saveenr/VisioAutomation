using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Shapes.Locking
{
    class LockCellsReader : SingleRowReader<LockCells>
    {
        public ColumnCell LockAspect { get; set; }
        public ColumnCell LockBegin { get; set; }
        public ColumnCell LockCalcWH { get; set; }
        public ColumnCell LockCrop { get; set; }
        public ColumnCell LockCustProp { get; set; }
        public ColumnCell LockDelete { get; set; }
        public ColumnCell LockEnd { get; set; }
        public ColumnCell LockFormat { get; set; }
        public ColumnCell LockFromGroupFormat { get; set; }
        public ColumnCell LockGroup { get; set; }
        public ColumnCell LockHeight { get; set; }
        public ColumnCell LockMoveX { get; set; }
        public ColumnCell LockMoveY { get; set; }
        public ColumnCell LockRotate { get; set; }
        public ColumnCell LockSelect { get; set; }
        public ColumnCell LockTextEdit { get; set; }
        public ColumnCell LockThemeColors { get; set; }
        public ColumnCell LockThemeEffects { get; set; }
        public ColumnCell LockVtxEdit { get; set; }
        public ColumnCell LockWidth { get; set; }

        public LockCellsReader()
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

        public override LockCells CellDataToCellGroup(ShapeSheet.CellData[] row)
        {
            var cells = new LockCells();
            cells.LockAspect = row[this.LockAspect];
            cells.LockBegin = row[this.LockBegin];
            cells.LockCalcWH = row[this.LockCalcWH];
            cells.LockCrop = row[this.LockCrop];
            cells.LockCustProp = row[this.LockCustProp];
            cells.LockDelete = row[this.LockDelete];
            cells.LockEnd = row[this.LockEnd];
            cells.LockFormat = row[this.LockFormat];
            cells.LockFromGroupFormat = row[this.LockFromGroupFormat];
            cells.LockGroup = row[this.LockGroup];
            cells.LockHeight = row[this.LockHeight];
            cells.LockMoveX = row[this.LockMoveX];
            cells.LockMoveY = row[this.LockMoveY];
            cells.LockRotate = row[this.LockRotate];
            cells.LockSelect = row[this.LockSelect];
            cells.LockTextEdit = row[this.LockTextEdit];
            cells.LockThemeColors = row[this.LockThemeColors];
            cells.LockThemeEffects = row[this.LockThemeEffects];
            cells.LockVtxEdit = row[this.LockVtxEdit];
            cells.LockWidth = row[this.LockWidth];
            return cells;
        }
    }
}