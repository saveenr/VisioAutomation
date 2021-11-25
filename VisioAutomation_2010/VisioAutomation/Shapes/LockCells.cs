using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class LockCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue Aspect { get; set; }
        public VisioAutomation.Core.CellValue Begin { get; set; }
        public VisioAutomation.Core.CellValue CalcWH { get; set; }
        public VisioAutomation.Core.CellValue Crop { get; set; }
        public VisioAutomation.Core.CellValue CustProp { get; set; }
        public VisioAutomation.Core.CellValue Delete { get; set; }
        public VisioAutomation.Core.CellValue End { get; set; }
        public VisioAutomation.Core.CellValue Format { get; set; }
        public VisioAutomation.Core.CellValue FromGroupFormat { get; set; }
        public VisioAutomation.Core.CellValue Group { get; set; }
        public VisioAutomation.Core.CellValue Height { get; set; }
        public VisioAutomation.Core.CellValue MoveX { get; set; }
        public VisioAutomation.Core.CellValue MoveY { get; set; }
        public VisioAutomation.Core.CellValue Rotate { get; set; }
        public VisioAutomation.Core.CellValue Select { get; set; }
        public VisioAutomation.Core.CellValue TextEdit { get; set; }
        public VisioAutomation.Core.CellValue ThemeColors { get; set; }
        public VisioAutomation.Core.CellValue ThemeEffects { get; set; }
        public VisioAutomation.Core.CellValue VertexEdit { get; set; }
        public VisioAutomation.Core.CellValue Width { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.Aspect), VisioAutomation.Core.SrcConstants.LockAspect, this.Aspect);
            yield return this.Create(nameof(this.Begin), VisioAutomation.Core.SrcConstants.LockBegin, this.Begin);
            yield return this.Create(nameof(this.CalcWH), VisioAutomation.Core.SrcConstants.LockCalcWH, this.CalcWH);
            yield return this.Create(nameof(this.Crop), VisioAutomation.Core.SrcConstants.LockCrop, this.Crop);
            yield return this.Create(nameof(this.CustProp), VisioAutomation.Core.SrcConstants.LockCustomProp, this.CustProp);
            yield return this.Create(nameof(this.Delete), VisioAutomation.Core.SrcConstants.LockDelete, this.Delete);
            yield return this.Create(nameof(this.End), VisioAutomation.Core.SrcConstants.LockEnd, this.End);
            yield return this.Create(nameof(this.Format), VisioAutomation.Core.SrcConstants.LockFormat, this.Format);
            yield return this.Create(nameof(this.FromGroupFormat), VisioAutomation.Core.SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
            yield return this.Create(nameof(this.Group), VisioAutomation.Core.SrcConstants.LockGroup, this.Group);
            yield return this.Create(nameof(this.Height), VisioAutomation.Core.SrcConstants.LockHeight, this.Height);
            yield return this.Create(nameof(this.MoveX), VisioAutomation.Core.SrcConstants.LockMoveX, this.MoveX);
            yield return this.Create(nameof(this.MoveY), VisioAutomation.Core.SrcConstants.LockMoveY, this.MoveY);
            yield return this.Create(nameof(this.Rotate), VisioAutomation.Core.SrcConstants.LockRotate, this.Rotate);
            yield return this.Create(nameof(this.Select), VisioAutomation.Core.SrcConstants.LockSelect, this.Select);
            yield return this.Create(nameof(this.TextEdit), VisioAutomation.Core.SrcConstants.LockTextEdit, this.TextEdit);
            yield return this.Create(nameof(this.ThemeColors), VisioAutomation.Core.SrcConstants.LockThemeColors, this.ThemeColors);
            yield return this.Create(nameof(this.ThemeEffects), VisioAutomation.Core.SrcConstants.LockThemeEffects, this.ThemeEffects);
            yield return this.Create(nameof(this.VertexEdit), VisioAutomation.Core.SrcConstants.LockVertexEdit, this.VertexEdit);
            yield return this.Create(nameof(this.Width), VisioAutomation.Core.SrcConstants.LockWidth, this.Width);
        }

        public static List<LockCells> GetCells(IVisio.Page page, IList<int> shapeid, VisioAutomation.Core.CellValueType type)
        {
            var reader = LockCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeid, type);
        }

        public static LockCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = LockCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<LockCellsBuilder> LockCells_lazy_builder = new System.Lazy<LockCellsBuilder>();


        class LockCellsBuilder : VASS.CellGroups.CellGroupBuilder<LockCells>
        {
            public LockCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override LockCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new LockCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

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