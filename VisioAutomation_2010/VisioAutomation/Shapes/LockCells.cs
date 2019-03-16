using System.Collections.Generic;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class LockCells : CellGroup
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

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.Aspect), SrcConstants.LockAspect, this.Aspect);
                yield return CellMetadataItem.Create(nameof(this.Begin), SrcConstants.LockBegin, this.Begin);
                yield return CellMetadataItem.Create(nameof(this.CalcWH), SrcConstants.LockCalcWH, this.CalcWH);
                yield return CellMetadataItem.Create(nameof(this.Crop), SrcConstants.LockCrop, this.Crop);
                yield return CellMetadataItem.Create(nameof(this.CustProp), SrcConstants.LockCustomProp, this.CustProp);
                yield return CellMetadataItem.Create(nameof(this.Delete), SrcConstants.LockDelete, this.Delete);
                yield return CellMetadataItem.Create(nameof(this.End), SrcConstants.LockEnd, this.End);
                yield return CellMetadataItem.Create(nameof(this.Format), SrcConstants.LockFormat, this.Format);
                yield return CellMetadataItem.Create(nameof(this.FromGroupFormat), SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
                yield return CellMetadataItem.Create(nameof(this.Group), SrcConstants.LockGroup, this.Group);
                yield return CellMetadataItem.Create(nameof(this.Height), SrcConstants.LockHeight, this.Height);
                yield return CellMetadataItem.Create(nameof(this.MoveX), SrcConstants.LockMoveX, this.MoveX);
                yield return CellMetadataItem.Create(nameof(this.MoveY), SrcConstants.LockMoveY, this.MoveY);
                yield return CellMetadataItem.Create(nameof(this.Rotate), SrcConstants.LockRotate, this.Rotate);
                yield return CellMetadataItem.Create(nameof(this.Select), SrcConstants.LockSelect, this.Select);
                yield return CellMetadataItem.Create(nameof(this.TextEdit), SrcConstants.LockTextEdit, this.TextEdit);
                yield return CellMetadataItem.Create(nameof(this.ThemeColors), SrcConstants.LockThemeColors, this.ThemeColors);
                yield return CellMetadataItem.Create(nameof(this.ThemeEffects), SrcConstants.LockThemeEffects, this.ThemeEffects);
                yield return CellMetadataItem.Create(nameof(this.VertexEdit), SrcConstants.LockVertexEdit, this.VertexEdit);
                yield return CellMetadataItem.Create(nameof(this.Width), SrcConstants.LockWidth, this.Width);
            }
        }

        public static List<LockCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = LockCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static LockCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = LockCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<LockCellsBuilder> LockCells_lazy_builder = new System.Lazy<LockCellsBuilder>();


        class LockCellsBuilder : CellGroupBuilder<LockCells>
        {
            public LockCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override LockCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new LockCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.Aspect = getcellvalue(nameof(LockCells.Aspect));
                cells.Begin = getcellvalue(nameof(LockCells.Begin));
                cells.CalcWH = getcellvalue(nameof(LockCells.CalcWH));
                cells.Crop = getcellvalue(nameof(LockCells.Crop));
                cells.CustProp = getcellvalue(nameof(LockCells.CustProp));
                cells.Delete = getcellvalue(nameof(LockCells.Delete));
                cells.End = getcellvalue(nameof(LockCells.End));
                cells.Format = getcellvalue(nameof(LockCells.Format));
                cells.FromGroupFormat = getcellvalue(nameof(LockCells.FromGroupFormat));
                cells.Group = getcellvalue(nameof(LockCells.Group));
                cells.Height = getcellvalue(nameof(LockCells.Height));
                cells.MoveX = getcellvalue(nameof(LockCells.MoveX));
                cells.MoveY = getcellvalue(nameof(LockCells.MoveY));
                cells.Rotate = getcellvalue(nameof(LockCells.Rotate));
                cells.Select = getcellvalue(nameof(LockCells.Select));
                cells.TextEdit = getcellvalue(nameof(LockCells.TextEdit));
                cells.ThemeColors = getcellvalue(nameof(LockCells.ThemeColors));
                cells.ThemeEffects = getcellvalue(nameof(LockCells.ThemeEffects));
                cells.VertexEdit = getcellvalue(nameof(LockCells.VertexEdit));
                cells.Width = getcellvalue(nameof(LockCells.Width));
                return cells;
            }
        }

    }
}