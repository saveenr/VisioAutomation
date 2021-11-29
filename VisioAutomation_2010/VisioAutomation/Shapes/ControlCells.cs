using System.Collections.Generic;
using VACG = VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ControlCells : VACG.CellGroup
    {
        public Core.CellValue CanGlue { get; set; }
        public Core.CellValue Tip { get; set; }
        public Core.CellValue X { get; set; }
        public Core.CellValue Y { get; set; }
        public Core.CellValue YBehavior { get; set; }
        public Core.CellValue XBehavior { get; set; }
        public Core.CellValue XDynamics { get; set; }
        public Core.CellValue YDynamics { get; set; }

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.CanGlue), Core.SrcConstants.ControlCanGlue, this.CanGlue);
            yield return this._create(nameof(this.Tip), Core.SrcConstants.ControlTip, this.Tip);
            yield return this._create(nameof(this.X), Core.SrcConstants.ControlX, this.X);
            yield return this._create(nameof(this.Y), Core.SrcConstants.ControlY, this.Y);
            yield return this._create(nameof(this.YBehavior), Core.SrcConstants.ControlYBehavior, this.YBehavior);
            yield return this._create(nameof(this.XBehavior), Core.SrcConstants.ControlXBehavior, this.XBehavior);
            yield return this._create(nameof(this.XDynamics), Core.SrcConstants.ControlXDynamics, this.XDynamics);
            yield return this._create(nameof(this.YDynamics), Core.SrcConstants.ControlYDynamics, this.YDynamics);
        }

        public static List<ControlCells> GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultiRow(shape, type);
        }

        public static List<List<ControlCells>> GetCells(IVisio.Page page, Core.ShapeIDPairs shapeidpairs, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultiRow(page, shapeidpairs, type);
        }


        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        class Builder : VACG.CellGroupBuilder<ControlCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.MultiRow)
            {
            }

            public override ControlCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new ControlCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);

                cells.CanGlue = getcellvalue(nameof(CanGlue));
                cells.Tip = getcellvalue(nameof(Tip));
                cells.X = getcellvalue(nameof(X));
                cells.Y = getcellvalue(nameof(Y));
                cells.YBehavior = getcellvalue(nameof(YBehavior));
                cells.XBehavior = getcellvalue(nameof(XBehavior));
                cells.XDynamics = getcellvalue(nameof(XDynamics));
                cells.YDynamics = getcellvalue(nameof(YDynamics));
                return cells;
            }
        }

    }
}