using System.Collections.Generic;
using VACG=VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class LockCells : VACG.CellGroup
    {
        public Core.CellValue Aspect { get; set; }
        public Core.CellValue Begin { get; set; }
        public Core.CellValue CalcWH { get; set; }
        public Core.CellValue Crop { get; set; }
        public Core.CellValue CustProp { get; set; }
        public Core.CellValue Delete { get; set; }
        public Core.CellValue End { get; set; }
        public Core.CellValue Format { get; set; }
        public Core.CellValue FromGroupFormat { get; set; }
        public Core.CellValue Group { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue MoveX { get; set; }
        public Core.CellValue MoveY { get; set; }
        public Core.CellValue Rotate { get; set; }
        public Core.CellValue Select { get; set; }
        public Core.CellValue TextEdit { get; set; }
        public Core.CellValue ThemeColors { get; set; }
        public Core.CellValue ThemeEffects { get; set; }
        public Core.CellValue VertexEdit { get; set; }
        public Core.CellValue Width { get; set; }

        public override IEnumerable<VACG.CellMetadataItem> GetCellMetadata()
        {
            yield return this._create(nameof(this.Aspect), Core.SrcConstants.LockAspect, this.Aspect);
            yield return this._create(nameof(this.Begin), Core.SrcConstants.LockBegin, this.Begin);
            yield return this._create(nameof(this.CalcWH), Core.SrcConstants.LockCalcWH, this.CalcWH);
            yield return this._create(nameof(this.Crop), Core.SrcConstants.LockCrop, this.Crop);
            yield return this._create(nameof(this.CustProp), Core.SrcConstants.LockCustomProp, this.CustProp);
            yield return this._create(nameof(this.Delete), Core.SrcConstants.LockDelete, this.Delete);
            yield return this._create(nameof(this.End), Core.SrcConstants.LockEnd, this.End);
            yield return this._create(nameof(this.Format), Core.SrcConstants.LockFormat, this.Format);
            yield return this._create(nameof(this.FromGroupFormat), Core.SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
            yield return this._create(nameof(this.Group), Core.SrcConstants.LockGroup, this.Group);
            yield return this._create(nameof(this.Height), Core.SrcConstants.LockHeight, this.Height);
            yield return this._create(nameof(this.MoveX), Core.SrcConstants.LockMoveX, this.MoveX);
            yield return this._create(nameof(this.MoveY), Core.SrcConstants.LockMoveY, this.MoveY);
            yield return this._create(nameof(this.Rotate), Core.SrcConstants.LockRotate, this.Rotate);
            yield return this._create(nameof(this.Select), Core.SrcConstants.LockSelect, this.Select);
            yield return this._create(nameof(this.TextEdit), Core.SrcConstants.LockTextEdit, this.TextEdit);
            yield return this._create(nameof(this.ThemeColors), Core.SrcConstants.LockThemeColors, this.ThemeColors);
            yield return this._create(nameof(this.ThemeEffects), Core.SrcConstants.LockThemeEffects, this.ThemeEffects);
            yield return this._create(nameof(this.VertexEdit), Core.SrcConstants.LockVertexEdit, this.VertexEdit);
            yield return this._create(nameof(this.Width), Core.SrcConstants.LockWidth, this.Width);
        }

        public static List<LockCells> GetCells(IVisio.Page page, IList<int> shapeid, Core.CellValueType type)
        {
            var reader = LockCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeid, type);
        }

        public static LockCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = LockCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> LockCells_lazy_builder = new System.Lazy<Builder>();


        class Builder : VACG.CellGroupBuilder<LockCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.SingleRow)
            {
            }

            public override LockCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new LockCells();
                var getcellvalue = row_to_cellgroup(row, cols);

                cells.Aspect = getcellvalue(nameof(Aspect));
                cells.Begin = getcellvalue(nameof(Begin));
                cells.CalcWH = getcellvalue(nameof(CalcWH));
                cells.Crop = getcellvalue(nameof(Crop));
                cells.CustProp = getcellvalue(nameof(CustProp));
                cells.Delete = getcellvalue(nameof(Delete));
                cells.End = getcellvalue(nameof(End));
                cells.Format = getcellvalue(nameof(Format));
                cells.FromGroupFormat = getcellvalue(nameof(FromGroupFormat));
                cells.Group = getcellvalue(nameof(Group));
                cells.Height = getcellvalue(nameof(Height));
                cells.MoveX = getcellvalue(nameof(MoveX));
                cells.MoveY = getcellvalue(nameof(MoveY));
                cells.Rotate = getcellvalue(nameof(Rotate));
                cells.Select = getcellvalue(nameof(Select));
                cells.TextEdit = getcellvalue(nameof(TextEdit));
                cells.ThemeColors = getcellvalue(nameof(ThemeColors));
                cells.ThemeEffects = getcellvalue(nameof(ThemeEffects));
                cells.VertexEdit = getcellvalue(nameof(VertexEdit));
                cells.Width = getcellvalue(nameof(Width));
                return cells;
            }
        }

    }
}