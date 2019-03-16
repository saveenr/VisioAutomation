using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ControlCells : CellGroup
    {
        public CellValueLiteral CanGlue { get; set; }
        public CellValueLiteral Tip { get; set; }
        public CellValueLiteral X { get; set; }
        public CellValueLiteral Y { get; set; }
        public CellValueLiteral YBehavior { get; set; }
        public CellValueLiteral XBehavior { get; set; }
        public CellValueLiteral XDynamics { get; set; }
        public CellValueLiteral YDynamics { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {
                yield return CellMetadataItem.Create(nameof(this.CanGlue), SrcConstants.ControlCanGlue, this.CanGlue);
                yield return CellMetadataItem.Create(nameof(this.Tip), SrcConstants.ControlTip, this.Tip);
                yield return CellMetadataItem.Create(nameof(this.X), SrcConstants.ControlX, this.X);
                yield return CellMetadataItem.Create(nameof(this.Y), SrcConstants.ControlY, this.Y);
                yield return CellMetadataItem.Create(nameof(this.YBehavior), SrcConstants.ControlYBehavior, this.YBehavior);
                yield return CellMetadataItem.Create(nameof(this.XBehavior), SrcConstants.ControlXBehavior, this.XBehavior);
                yield return CellMetadataItem.Create(nameof(this.XDynamics), SrcConstants.ControlXDynamics, this.XDynamics);
                yield return CellMetadataItem.Create(nameof(this.YDynamics), SrcConstants.ControlYDynamics, this.YDynamics);
            }
        }

        public static List<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ControlCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeids, type);
        }

        public static List<ControlCells> GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ControlCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        private static readonly System.Lazy<ControlCellsBuilder> ControlCells_lazy_builder = new System.Lazy<ControlCellsBuilder>();

        class ControlCellsBuilder : CellGroupBuilder<ControlCells>
        {
            public ControlCellsBuilder() : base(CellGroupBuilderType.MultiRow)
            {
            }

            public override ControlCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new ControlCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

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