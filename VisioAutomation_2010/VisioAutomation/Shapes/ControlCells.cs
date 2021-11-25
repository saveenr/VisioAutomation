using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ControlCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue CanGlue { get; set; }
        public VisioAutomation.Core.CellValue Tip { get; set; }
        public VisioAutomation.Core.CellValue X { get; set; }
        public VisioAutomation.Core.CellValue Y { get; set; }
        public VisioAutomation.Core.CellValue YBehavior { get; set; }
        public VisioAutomation.Core.CellValue XBehavior { get; set; }
        public VisioAutomation.Core.CellValue XDynamics { get; set; }
        public VisioAutomation.Core.CellValue YDynamics { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.CanGlue), VisioAutomation.Core.SrcConstants.ControlCanGlue, this.CanGlue);
            yield return this.Create(nameof(this.Tip), VisioAutomation.Core.SrcConstants.ControlTip, this.Tip);
            yield return this.Create(nameof(this.X), VisioAutomation.Core.SrcConstants.ControlX, this.X);
            yield return this.Create(nameof(this.Y), VisioAutomation.Core.SrcConstants.ControlY, this.Y);
            yield return this.Create(nameof(this.YBehavior), VisioAutomation.Core.SrcConstants.ControlYBehavior, this.YBehavior);
            yield return this.Create(nameof(this.XBehavior), VisioAutomation.Core.SrcConstants.ControlXBehavior, this.XBehavior);
            yield return this.Create(nameof(this.XDynamics), VisioAutomation.Core.SrcConstants.ControlXDynamics, this.XDynamics);
            yield return this.Create(nameof(this.YDynamics), VisioAutomation.Core.SrcConstants.ControlYDynamics, this.YDynamics);
        }

        public static List<ControlCells> GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = ControlCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        public static List<List<ControlCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, VisioAutomation.Core.CellValueType type)
        {
            var reader = ControlCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }


        private static readonly System.Lazy<ControlCellsBuilder> ControlCells_lazy_builder = new System.Lazy<ControlCellsBuilder>();

        class ControlCellsBuilder : VASS.CellGroups.CellGroupBuilder<ControlCells>
        {
            public ControlCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.MultiRow)
            {
            }

            public override ControlCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new ControlCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.CanGlue = getcellvalue(nameof(ControlCells.CanGlue));
                cells.Tip = getcellvalue(nameof(ControlCells.Tip));
                cells.X = getcellvalue(nameof(ControlCells.X));
                cells.Y = getcellvalue(nameof(ControlCells.Y));
                cells.YBehavior = getcellvalue(nameof(ControlCells.YBehavior));
                cells.XBehavior = getcellvalue(nameof(ControlCells.XBehavior));
                cells.XDynamics = getcellvalue(nameof(ControlCells.XDynamics));
                cells.YDynamics = getcellvalue(nameof(ControlCells.YDynamics));
                return cells;
            }
        }

    }
}