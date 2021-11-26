using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;

using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class XFormCells : CellGroup
    {
        public Core.CellValue PinX { get; set; }
        public Core.CellValue PinY { get; set; }
        public Core.CellValue LocPinX { get; set; }
        public Core.CellValue LocPinY { get; set; }
        public Core.CellValue Width { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue Angle { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.PinX), Core.SrcConstants.XFormPinX, this.PinX);
            yield return this.Create(nameof(this.PinY), Core.SrcConstants.XFormPinY, this.PinY);
            yield return this.Create(nameof(this.LocPinX), Core.SrcConstants.XFormLocPinX, this.LocPinX);
            yield return this.Create(nameof(this.LocPinY), Core.SrcConstants.XFormLocPinY, this.LocPinY);
            yield return this.Create(nameof(this.Width), Core.SrcConstants.XFormWidth, this.Width);
            yield return this.Create(nameof(this.Height), Core.SrcConstants.XFormHeight, this.Height);
            yield return this.Create(nameof(this.Angle), Core.SrcConstants.XFormAngle, this.Angle);
        }


        public static List<XFormCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static XFormCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> ShapeXFormCells_lazy_builder = new System.Lazy<Builder>();

        class Builder : CellGroupBuilder<XFormCells>
        {
            public Builder() : base(CellGroupBuilderType.SingleRow)
            {
            }

            public override XFormCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new XFormCells();
                var getcellvalue = row_to_cellgroup(row, cols);

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