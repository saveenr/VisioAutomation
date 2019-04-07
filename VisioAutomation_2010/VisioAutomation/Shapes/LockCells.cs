using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class LockCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue Aspect { get; set; }
        public VASS.CellValue Begin { get; set; }
        public VASS.CellValue CalcWH { get; set; }
        public VASS.CellValue Crop { get; set; }
        public VASS.CellValue CustProp { get; set; }
        public VASS.CellValue Delete { get; set; }
        public VASS.CellValue End { get; set; }
        public VASS.CellValue Format { get; set; }
        public VASS.CellValue FromGroupFormat { get; set; }
        public VASS.CellValue Group { get; set; }
        public VASS.CellValue Height { get; set; }
        public VASS.CellValue MoveX { get; set; }
        public VASS.CellValue MoveY { get; set; }
        public VASS.CellValue Rotate { get; set; }
        public VASS.CellValue Select { get; set; }
        public VASS.CellValue TextEdit { get; set; }
        public VASS.CellValue ThemeColors { get; set; }
        public VASS.CellValue ThemeEffects { get; set; }
        public VASS.CellValue VertexEdit { get; set; }
        public VASS.CellValue Width { get; set; }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {


                yield return this.Create(nameof(this.Aspect), VASS.SrcConstants.LockAspect, this.Aspect);
                yield return this.Create(nameof(this.Begin), VASS.SrcConstants.LockBegin, this.Begin);
                yield return this.Create(nameof(this.CalcWH), VASS.SrcConstants.LockCalcWH, this.CalcWH);
                yield return this.Create(nameof(this.Crop), VASS.SrcConstants.LockCrop, this.Crop);
                yield return this.Create(nameof(this.CustProp), VASS.SrcConstants.LockCustomProp, this.CustProp);
                yield return this.Create(nameof(this.Delete), VASS.SrcConstants.LockDelete, this.Delete);
                yield return this.Create(nameof(this.End), VASS.SrcConstants.LockEnd, this.End);
                yield return this.Create(nameof(this.Format), VASS.SrcConstants.LockFormat, this.Format);
                yield return this.Create(nameof(this.FromGroupFormat), VASS.SrcConstants.LockFromGroupFormat, this.FromGroupFormat);
                yield return this.Create(nameof(this.Group), VASS.SrcConstants.LockGroup, this.Group);
                yield return this.Create(nameof(this.Height), VASS.SrcConstants.LockHeight, this.Height);
                yield return this.Create(nameof(this.MoveX), VASS.SrcConstants.LockMoveX, this.MoveX);
                yield return this.Create(nameof(this.MoveY), VASS.SrcConstants.LockMoveY, this.MoveY);
                yield return this.Create(nameof(this.Rotate), VASS.SrcConstants.LockRotate, this.Rotate);
                yield return this.Create(nameof(this.Select), VASS.SrcConstants.LockSelect, this.Select);
                yield return this.Create(nameof(this.TextEdit), VASS.SrcConstants.LockTextEdit, this.TextEdit);
                yield return this.Create(nameof(this.ThemeColors), VASS.SrcConstants.LockThemeColors, this.ThemeColors);
                yield return this.Create(nameof(this.ThemeEffects), VASS.SrcConstants.LockThemeEffects, this.ThemeEffects);
                yield return this.Create(nameof(this.VertexEdit), VASS.SrcConstants.LockVertexEdit, this.VertexEdit);
                yield return this.Create(nameof(this.Width), VASS.SrcConstants.LockWidth, this.Width);
            }
        }

        public static List<LockCells> GetCells(IVisio.Page page, IList<int> shapeid, VASS.CellValueType type)
        {
            var reader = LockCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeid, type);
        }

        public static LockCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
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