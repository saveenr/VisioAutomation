using System.Collections.Generic;
using VASS=VisioAutomation.ShapeSheet;

using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue PinX { get; set; }
        public VASS.CellValue PinY { get; set; }
        public VASS.CellValue LocPinX { get; set; }
        public VASS.CellValue LocPinY { get; set; }
        public VASS.CellValue Width { get; set; }
        public VASS.CellValue Height { get; set; }
        public VASS.CellValue Angle { get; set; }

        public override IEnumerable<VASS.CellGroups.CellMetadataItem> CellMetadata
        {
            get
            {


                yield return this.Create(nameof(this.PinX), VASS.SrcConstants.XFormPinX, this.PinX);
                yield return this.Create(nameof(this.PinY), VASS.SrcConstants.XFormPinY, this.PinY);
                yield return this.Create(nameof(this.LocPinX), VASS.SrcConstants.XFormLocPinX, this.LocPinX);
                yield return this.Create(nameof(this.LocPinY), VASS.SrcConstants.XFormLocPinY, this.LocPinY);
                yield return this.Create(nameof(this.Width), VASS.SrcConstants.XFormWidth, this.Width);
                yield return this.Create(nameof(this.Height), VASS.SrcConstants.XFormHeight, this.Height);
                yield return this.Create(nameof(this.Angle), VASS.SrcConstants.XFormAngle, this.Angle);
            }
        }


        public static List<ShapeXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, VASS.CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeXFormCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeXFormCellsBuilder> ShapeXFormCells_lazy_builder = new System.Lazy<ShapeXFormCellsBuilder>();

        class ShapeXFormCellsBuilder : VASS.CellGroups.CellGroupBuilder<ShapeXFormCells>
        {
            public ShapeXFormCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override ShapeXFormCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new ShapeXFormCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.PinX = getcellvalue(nameof(ShapeXFormCells.PinX));
                cells.PinY = getcellvalue(nameof(ShapeXFormCells.PinY));
                cells.LocPinX = getcellvalue(nameof(ShapeXFormCells.LocPinX));
                cells.LocPinY = getcellvalue(nameof(ShapeXFormCells.LocPinY));
                cells.Width = getcellvalue(nameof(ShapeXFormCells.Width));
                cells.Height = getcellvalue(nameof(ShapeXFormCells.Height));
                cells.Angle = getcellvalue(nameof(ShapeXFormCells.Angle));

                return cells;
            }
        }

    }
}