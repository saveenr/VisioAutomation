using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;



namespace VisioAutomation.Shapes
{
    public static class LockHelper
    {
        public static List<LockCells> GetLockCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = LockCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static LockCells GetLockCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = LockCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<LockCellsReader> LockCells_lazy_reader = new System.Lazy<LockCellsReader>();


        class LockCellsReader : CellGroupReader<LockCells>
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

            public LockCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.Aspect = this.query_singlerow.Columns.Add(SrcConstants.LockAspect, nameof(this.Aspect));
                this.Begin = this.query_singlerow.Columns.Add(SrcConstants.LockBegin, nameof(this.Begin));
                this.CalcWH = this.query_singlerow.Columns.Add(SrcConstants.LockCalcWH, nameof(this.CalcWH));
                this.Crop = this.query_singlerow.Columns.Add(SrcConstants.LockCrop, nameof(this.Crop));
                this.CustomProp = this.query_singlerow.Columns.Add(SrcConstants.LockCustomProp, nameof(this.CustomProp));
                this.Delete = this.query_singlerow.Columns.Add(SrcConstants.LockDelete, nameof(this.Delete));
                this.End = this.query_singlerow.Columns.Add(SrcConstants.LockEnd, nameof(this.End));
                this.Format = this.query_singlerow.Columns.Add(SrcConstants.LockFormat, nameof(this.Format));
                this.FromGroupFormat = this.query_singlerow.Columns.Add(SrcConstants.LockFromGroupFormat, nameof(this.FromGroupFormat));
                this.Group = this.query_singlerow.Columns.Add(SrcConstants.LockGroup, nameof(this.Group));
                this.Height = this.query_singlerow.Columns.Add(SrcConstants.LockHeight, nameof(this.Height));
                this.MoveX = this.query_singlerow.Columns.Add(SrcConstants.LockMoveX, nameof(this.MoveX));
                this.MoveY = this.query_singlerow.Columns.Add(SrcConstants.LockMoveY, nameof(this.MoveY));
                this.Rotate = this.query_singlerow.Columns.Add(SrcConstants.LockRotate, nameof(this.Rotate));
                this.Select = this.query_singlerow.Columns.Add(SrcConstants.LockSelect, nameof(this.Select));
                this.TextEdit = this.query_singlerow.Columns.Add(SrcConstants.LockTextEdit, nameof(this.TextEdit));
                this.ThemeColors = this.query_singlerow.Columns.Add(SrcConstants.LockThemeColors, nameof(this.ThemeColors));
                this.ThemeEffects = this.query_singlerow.Columns.Add(SrcConstants.LockThemeEffects, nameof(this.ThemeEffects));
                this.VertexEdit = this.query_singlerow.Columns.Add(SrcConstants.LockVertexEdit, nameof(this.VertexEdit));
                this.Width = this.query_singlerow.Columns.Add(SrcConstants.LockWidth, nameof(this.Width));
            }

            public override LockCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
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
}