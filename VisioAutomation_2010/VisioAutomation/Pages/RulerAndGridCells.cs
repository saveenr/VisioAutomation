using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class RulerAndGridCells : CellGroup
    {
        public Core.CellValue XGridDensity { get; set; }
        public Core.CellValue YGridDensity { get; set; }
        public Core.CellValue XGridOrigin { get; set; }
        public Core.CellValue YGridOrigin { get; set; }
        public Core.CellValue XGridSpacing { get; set; }
        public Core.CellValue YGridSpacing { get; set; }
        public Core.CellValue XRulerDensity { get; set; }
        public Core.CellValue XRulerOrigin { get; set; }
        public Core.CellValue YRulerDensity { get; set; }
        public Core.CellValue YRulerOrigin { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this._create(nameof(this.XGridDensity), Core.SrcConstants.XGridDensity, this.XGridDensity);
            yield return this._create(nameof(this.XGridOrigin), Core.SrcConstants.XGridOrigin, this.XGridOrigin);
            yield return this._create(nameof(this.XGridSpacing), Core.SrcConstants.XGridSpacing, this.XGridSpacing);
            yield return this._create(nameof(this.XRulerDensity), Core.SrcConstants.XRulerDensity, this.XRulerDensity);
            yield return this._create(nameof(this.XRulerOrigin), Core.SrcConstants.XRulerOrigin, this.XRulerOrigin);
            yield return this._create(nameof(this.YGridDensity), Core.SrcConstants.YGridDensity, this.YGridDensity);
            yield return this._create(nameof(this.YGridOrigin), Core.SrcConstants.YGridOrigin, this.YGridOrigin);
            yield return this._create(nameof(this.YGridSpacing), Core.SrcConstants.YGridSpacing, this.YGridSpacing);
            yield return this._create(nameof(this.YRulerDensity), Core.SrcConstants.YRulerDensity, this.YRulerDensity);
            yield return this._create(nameof(this.YRulerOrigin), Core.SrcConstants.YRulerOrigin, this.YRulerOrigin);
        }

        public static RulerAndGridCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = PageRulerAndGridCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> PageRulerAndGridCells_lazy_builder = new System.Lazy<Builder>();

        class Builder : CellGroupBuilder<RulerAndGridCells>
        {
            public Builder() : base(CellGroupBuilderType.SingleRow)
            {
            }

            public override RulerAndGridCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new RulerAndGridCells();
                var getcellvalue = row_to_cellgroup(row, cols);

                cells.XGridDensity = getcellvalue(nameof(XGridDensity));
                cells.XGridOrigin = getcellvalue(nameof(XGridOrigin));
                cells.XGridSpacing = getcellvalue(nameof(XGridSpacing));
                cells.XRulerDensity = getcellvalue(nameof(XRulerDensity));
                cells.XRulerOrigin = getcellvalue(nameof(XRulerOrigin));
                cells.YGridDensity = getcellvalue(nameof(YGridDensity));
                cells.YGridOrigin = getcellvalue(nameof(YGridOrigin));
                cells.YGridSpacing = getcellvalue(nameof(YGridSpacing));
                cells.YRulerDensity = getcellvalue(nameof(YRulerDensity));
                cells.YRulerOrigin = getcellvalue(nameof(YRulerOrigin));

                return cells;
            }
        }

    }
}