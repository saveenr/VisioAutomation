using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageRulerAndGridCells : CellRecord
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

        public override IEnumerable<CellMetadata> GetCellMetadata()
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

        public static PageRulerAndGridCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static PageRulerAndGridCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var record = new PageRulerAndGridCells();
            var getcellvalue = getvalfromrowfunc(row, cols);

            record.XGridDensity = getcellvalue(nameof(XGridDensity));
            record.XGridOrigin = getcellvalue(nameof(XGridOrigin));
            record.XGridSpacing = getcellvalue(nameof(XGridSpacing));
            record.XRulerDensity = getcellvalue(nameof(XRulerDensity));
            record.XRulerOrigin = getcellvalue(nameof(XRulerOrigin));
            record.YGridDensity = getcellvalue(nameof(YGridDensity));
            record.YGridOrigin = getcellvalue(nameof(YGridOrigin));
            record.YGridSpacing = getcellvalue(nameof(YGridSpacing));
            record.YRulerDensity = getcellvalue(nameof(YRulerDensity));
            record.YRulerOrigin = getcellvalue(nameof(YRulerOrigin));

            return record;
        }

        class Builder : CellRecordBuilderCellQuery<PageRulerAndGridCells>
        {
            public Builder() : base(PageRulerAndGridCells.RowToRecord)
            {
            }

        }
    }
}