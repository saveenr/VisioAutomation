using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    class LockCellsReader : ReaderSingleRow<LockCells>
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
            this.Aspect = this.query.AddColumn(SrcConstants.LockAspect, nameof(SrcConstants.LockAspect));
            this.Begin = this.query.AddColumn(SrcConstants.LockBegin, nameof(SrcConstants.LockBegin));
            this.CalcWH = this.query.AddColumn(SrcConstants.LockCalcWH, nameof(SrcConstants.LockCalcWH));
            this.Crop = this.query.AddColumn(SrcConstants.LockCrop, nameof(SrcConstants.LockCrop));
            this.CustomProp = this.query.AddColumn(SrcConstants.LockCustomProp, nameof(SrcConstants.LockCustomProp));
            this.Delete = this.query.AddColumn(SrcConstants.LockDelete, nameof(SrcConstants.LockDelete));
            this.End = this.query.AddColumn(SrcConstants.LockEnd, nameof(SrcConstants.LockEnd));
            this.Format = this.query.AddColumn(SrcConstants.LockFormat, nameof(SrcConstants.LockFormat));
            this.FromGroupFormat = this.query.AddColumn(SrcConstants.LockFromGroupFormat, nameof(SrcConstants.LockFromGroupFormat));
            this.Group = this.query.AddColumn(SrcConstants.LockGroup, nameof(SrcConstants.LockGroup));
            this.Height = this.query.AddColumn(SrcConstants.LockHeight, nameof(SrcConstants.LockHeight));
            this.MoveX = this.query.AddColumn(SrcConstants.LockMoveX, nameof(SrcConstants.LockMoveX));
            this.MoveY = this.query.AddColumn(SrcConstants.LockMoveY, nameof(SrcConstants.LockMoveY));
            this.Rotate = this.query.AddColumn(SrcConstants.LockRotate, nameof(SrcConstants.LockRotate));
            this.Select = this.query.AddColumn(SrcConstants.LockSelect, nameof(SrcConstants.LockSelect));
            this.TextEdit = this.query.AddColumn(SrcConstants.LockTextEdit, nameof(SrcConstants.LockTextEdit));
            this.ThemeColors = this.query.AddColumn(SrcConstants.LockThemeColors, nameof(SrcConstants.LockThemeColors));
            this.ThemeEffects = this.query.AddColumn(SrcConstants.LockThemeEffects, nameof(SrcConstants.LockThemeEffects));
            this.VertexEdit = this.query.AddColumn(SrcConstants.LockVertexEdit, nameof(SrcConstants.LockVertexEdit));
            this.Width = this.query.AddColumn(SrcConstants.LockWidth, nameof(SrcConstants.LockWidth));
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