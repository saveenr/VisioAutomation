using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class LockCells : CellGroupBase
    {
        public CellValueLiteral Aspect { get; set; }
        public CellValueLiteral Begin { get; set; }
        public CellValueLiteral CalcWH { get; set; }
        public CellValueLiteral Crop { get; set; }
        public CellValueLiteral CustProp { get; set; }
        public CellValueLiteral Delete { get; set; }
        public CellValueLiteral End { get; set; }
        public CellValueLiteral Format { get; set; }
        public CellValueLiteral FromGroupFormat { get; set; }
        public CellValueLiteral Group { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral MoveX { get; set; }
        public CellValueLiteral MoveY { get; set; }
        public CellValueLiteral Rotate { get; set; }
        public CellValueLiteral Select { get; set; }
        public CellValueLiteral TextEdit { get; set; }
        public CellValueLiteral ThemeColors { get; set; }
        public CellValueLiteral ThemeEffects { get; set; }
        public CellValueLiteral VertexEdit { get; set; }
        public CellValueLiteral Width { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.LockAspect, this.Aspect);
                yield return SrcValuePair.Create(SrcConstants.LockBegin, this.Begin);
                yield return SrcValuePair.Create(SrcConstants.LockCalcWH, this.CalcWH);
                yield return SrcValuePair.Create(SrcConstants.LockCrop, this.Crop);
                yield return SrcValuePair.Create(SrcConstants.LockCustomProp, this.CustProp);
                yield return SrcValuePair.Create(SrcConstants.LockDelete, this.Delete);
                yield return SrcValuePair.Create(SrcConstants.LockEnd, this.End);
                yield return SrcValuePair.Create(SrcConstants.LockFormat, this.Format);
                yield return SrcValuePair.Create(SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
                yield return SrcValuePair.Create(SrcConstants.LockGroup, this.Group);
                yield return SrcValuePair.Create(SrcConstants.LockHeight, this.Height);
                yield return SrcValuePair.Create(SrcConstants.LockMoveX, this.MoveX);
                yield return SrcValuePair.Create(SrcConstants.LockMoveY, this.MoveY);
                yield return SrcValuePair.Create(SrcConstants.LockRotate, this.Rotate);
                yield return SrcValuePair.Create(SrcConstants.LockSelect, this.Select);
                yield return SrcValuePair.Create(SrcConstants.LockTextEdit, this.TextEdit);
                yield return SrcValuePair.Create(SrcConstants.LockThemeColors, this.ThemeColors);
                yield return SrcValuePair.Create(SrcConstants.LockThemeEffects, this.ThemeEffects);
                yield return SrcValuePair.Create(SrcConstants.LockVertexEdit, this.VertexEdit);
                yield return SrcValuePair.Create(SrcConstants.LockWidth, this.Width);
            }
        }


        public static List<LockCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static LockCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<LockCellsReader> lazy_reader = new System.Lazy<LockCellsReader>();


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