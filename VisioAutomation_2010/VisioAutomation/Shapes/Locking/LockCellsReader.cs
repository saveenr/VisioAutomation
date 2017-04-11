using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Locking
{
    class LockCellsReader : SingleRowReader<LockCells>
    {
        public CellColumn Aspect { get; set; }
        public CellColumn Begin { get; set; }
        public CellColumn CalcWH { get; set; }
        public CellColumn Crop { get; set; }
        public CellColumn CustomProp { get; set; }
        public CellColumn Delete { get; set; }
        public CellColumn End { get; set; }
        public CellColumn Format { get; set; }
        public CellColumn FromGroupFormat { get; set; }
        public CellColumn Group { get; set; }
        public CellColumn Height { get; set; }
        public CellColumn MoveX { get; set; }
        public CellColumn MoveY { get; set; }
        public CellColumn Rotate { get; set; }
        public CellColumn Select { get; set; }
        public CellColumn TextEdit { get; set; }
        public CellColumn ThemeColors { get; set; }
        public CellColumn ThemeEffects { get; set; }
        public CellColumn VertexEdit { get; set; }
        public CellColumn Width { get; set; }

        public LockCellsReader()
        {
            this.Aspect = this.query.AddCell(SrcConstants.LockAspect, nameof(SrcConstants.LockAspect));
            this.Begin = this.query.AddCell(SrcConstants.LockBegin, nameof(SrcConstants.LockBegin));
            this.CalcWH = this.query.AddCell(SrcConstants.LockCalcWH, nameof(SrcConstants.LockCalcWH));
            this.Crop = this.query.AddCell(SrcConstants.LockCrop, nameof(SrcConstants.LockCrop));
            this.CustomProp = this.query.AddCell(SrcConstants.LockCustomProp, nameof(SrcConstants.LockCustomProp));
            this.Delete = this.query.AddCell(SrcConstants.LockDelete, nameof(SrcConstants.LockDelete));
            this.End = this.query.AddCell(SrcConstants.LockEnd, nameof(SrcConstants.LockEnd));
            this.Format = this.query.AddCell(SrcConstants.LockFormat, nameof(SrcConstants.LockFormat));
            this.FromGroupFormat = this.query.AddCell(SrcConstants.LockFromGroupFormat, nameof(SrcConstants.LockFromGroupFormat));
            this.Group = this.query.AddCell(SrcConstants.LockGroup, nameof(SrcConstants.LockGroup));
            this.Height = this.query.AddCell(SrcConstants.LockHeight, nameof(SrcConstants.LockHeight));
            this.MoveX = this.query.AddCell(SrcConstants.LockMoveX, nameof(SrcConstants.LockMoveX));
            this.MoveY = this.query.AddCell(SrcConstants.LockMoveY, nameof(SrcConstants.LockMoveY));
            this.Rotate = this.query.AddCell(SrcConstants.LockRotate, nameof(SrcConstants.LockRotate));
            this.Select = this.query.AddCell(SrcConstants.LockSelect, nameof(SrcConstants.LockSelect));
            this.TextEdit = this.query.AddCell(SrcConstants.LockTextEdit, nameof(SrcConstants.LockTextEdit));
            this.ThemeColors = this.query.AddCell(SrcConstants.LockThemeColors, nameof(SrcConstants.LockThemeColors));
            this.ThemeEffects = this.query.AddCell(SrcConstants.LockThemeEffects, nameof(SrcConstants.LockThemeEffects));
            this.VertexEdit = this.query.AddCell(SrcConstants.LockVertexEdit, nameof(SrcConstants.LockVertexEdit));
            this.Width = this.query.AddCell(SrcConstants.LockWidth, nameof(SrcConstants.LockWidth));
        }

        public override LockCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new LockCells();
            cells.Aspect = row[this.Aspect];
            cells.Begin = row[this.Begin];
            cells.CalcWH = row[this.CalcWH];
            cells.Crop = row[this.Crop];
            cells.CustProp = row[this.CustomProp];
            cells.Delete = row[this.Delete];
            cells.End = row[this.End];
            cells.Format = row[this.Format];
            cells.FromGroupFormat = row[this.FromGroupFormat];
            cells.Group = row[this.Group];
            cells.Height = row[this.Height];
            cells.MoveX = row[this.MoveX];
            cells.MoveY = row[this.MoveY];
            cells.Rotate = row[this.Rotate];
            cells.Select = row[this.Select];
            cells.TextEdit = row[this.TextEdit];
            cells.ThemeColors = row[this.ThemeColors];
            cells.ThemeEffects = row[this.ThemeEffects];
            cells.VertexEdit = row[this.VertexEdit];
            cells.Width = row[this.Width];
            return cells;
        }
    }
}