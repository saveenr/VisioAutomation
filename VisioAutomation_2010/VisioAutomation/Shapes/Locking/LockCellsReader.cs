using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Locking
{
    class LockCellsReader : SingleRowReader<LockCells>
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

        public LockCellsReader()
        {
            this.LockAspect = this.query.AddCell(SrcConstants.LockAspect, nameof(SrcConstants.LockAspect));
            this.LockBegin = this.query.AddCell(SrcConstants.LockBegin, nameof(SrcConstants.LockBegin));
            this.LockCalcWH = this.query.AddCell(SrcConstants.LockCalcWH, nameof(SrcConstants.LockCalcWH));
            this.LockCrop = this.query.AddCell(SrcConstants.LockCrop, nameof(SrcConstants.LockCrop));
            this.LockCustProp = this.query.AddCell(SrcConstants.LockCustProp, nameof(SrcConstants.LockCustProp));
            this.LockDelete = this.query.AddCell(SrcConstants.LockDelete, nameof(SrcConstants.LockDelete));
            this.LockEnd = this.query.AddCell(SrcConstants.LockEnd, nameof(SrcConstants.LockEnd));
            this.LockFormat = this.query.AddCell(SrcConstants.LockFormat, nameof(SrcConstants.LockFormat));
            this.LockFromGroupFormat = this.query.AddCell(SrcConstants.LockFromGroupFormat, nameof(SrcConstants.LockFromGroupFormat));
            this.LockGroup = this.query.AddCell(SrcConstants.LockGroup, nameof(SrcConstants.LockGroup));
            this.LockHeight = this.query.AddCell(SrcConstants.LockHeight, nameof(SrcConstants.LockHeight));
            this.LockMoveX = this.query.AddCell(SrcConstants.LockMoveX, nameof(SrcConstants.LockMoveX));
            this.LockMoveY = this.query.AddCell(SrcConstants.LockMoveY, nameof(SrcConstants.LockMoveY));
            this.LockRotate = this.query.AddCell(SrcConstants.LockRotate, nameof(SrcConstants.LockRotate));
            this.LockSelect = this.query.AddCell(SrcConstants.LockSelect, nameof(SrcConstants.LockSelect));
            this.LockTextEdit = this.query.AddCell(SrcConstants.LockTextEdit, nameof(SrcConstants.LockTextEdit));
            this.LockThemeColors = this.query.AddCell(SrcConstants.LockThemeColors, nameof(SrcConstants.LockThemeColors));
            this.LockThemeEffects = this.query.AddCell(SrcConstants.LockThemeEffects, nameof(SrcConstants.LockThemeEffects));
            this.LockVtxEdit = this.query.AddCell(SrcConstants.LockVertexEdit, nameof(SrcConstants.LockVertexEdit));
            this.LockWidth = this.query.AddCell(SrcConstants.LockWidth, nameof(SrcConstants.LockWidth));
        }

        public override LockCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
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