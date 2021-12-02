using System.Collections.Generic;
using VACG = VisioAutomation.ShapeSheet.CellGroups;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class XFormCells : VACG.CellGroup
    {
        public Core.CellValue PinX { get; set; }
        public Core.CellValue PinY { get; set; }
        public Core.CellValue LocPinX { get; set; }
        public Core.CellValue LocPinY { get; set; }
        public Core.CellValue Width { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue Angle { get; set; }

        public override IEnumerable<VACG.CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.PinX), Core.SrcConstants.XFormPinX, this.PinX);
            yield return this._create(nameof(this.PinY), Core.SrcConstants.XFormPinY, this.PinY);
            yield return this._create(nameof(this.LocPinX), Core.SrcConstants.XFormLocPinX, this.LocPinX);
            yield return this._create(nameof(this.LocPinY), Core.SrcConstants.XFormLocPinY, this.LocPinY);
            yield return this._create(nameof(this.Width), Core.SrcConstants.XFormWidth, this.Width);
            yield return this._create(nameof(this.Height), Core.SrcConstants.XFormHeight, this.Height);
            yield return this._create(nameof(this.Angle), Core.SrcConstants.XFormAngle, this.Angle);
        }


        public static List<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesSingleRow(page, shapeids, type);
        }

        public static XFormCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        class Builder : VACG.CellGroupBuilder<XFormCells>
        {
            public Builder() : base(VACG.CellGroupBuilderType.SingleRow)
            {
            }

            public override XFormCells ToCellGroup(VASS.Data.CellValueRow<string> row, VASS.Query.Columns cols)
            {
                var cells = new XFormCells();
                var getcellvalue = queryrow_to_cellgroup(row, cols);

                cells.PinX = getcellvalue(nameof(PinX));
                cells.PinY = getcellvalue(nameof(PinY));
                cells.LocPinX = getcellvalue(nameof(LocPinX));
                cells.LocPinY = getcellvalue(nameof(LocPinY));
                cells.Width = getcellvalue(nameof(Width));
                cells.Height = getcellvalue(nameof(Height));
                cells.Angle = getcellvalue(nameof(Angle));

                return cells;
            }
        }
    }
}