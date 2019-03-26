using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ControlCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValueLiteral CanGlue { get; set; }
        public VASS.CellValueLiteral Tip { get; set; }
        public VASS.CellValueLiteral X { get; set; }
        public VASS.CellValueLiteral Y { get; set; }
        public VASS.CellValueLiteral YBehavior { get; set; }
        public VASS.CellValueLiteral XBehavior { get; set; }
        public VASS.CellValueLiteral XDynamics { get; set; }
        public VASS.CellValueLiteral YDynamics { get; set; }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {
                yield return this.Create(nameof(this.CanGlue), VASS.SrcConstants.ControlCanGlue, this.CanGlue);
                yield return this.Create(nameof(this.Tip), VASS.SrcConstants.ControlTip, this.Tip);
                yield return this.Create(nameof(this.X), VASS.SrcConstants.ControlX, this.X);
                yield return this.Create(nameof(this.Y), VASS.SrcConstants.ControlY, this.Y);
                yield return this.Create(nameof(this.YBehavior), VASS.SrcConstants.ControlYBehavior, this.YBehavior);
                yield return this.Create(nameof(this.XBehavior), VASS.SrcConstants.ControlXBehavior, this.XBehavior);
                yield return this.Create(nameof(this.XDynamics), VASS.SrcConstants.ControlXDynamics, this.XDynamics);
                yield return this.Create(nameof(this.YDynamics), VASS.SrcConstants.ControlYDynamics, this.YDynamics);
            }
        }

        public static List<List<ControlCells>> GetCells(IVisio.Page page, VASS.Query.ShapeIdPairs shapeidpairs, VASS.CellValueType type)
        {
            var reader = ControlCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }

        public static List<ControlCells> GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = ControlCells_lazy_builder.Value;
            return reader.GetCellsMultiRow(shape, type);
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