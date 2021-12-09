using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class PageFormatCells : CellRecord
    {
        public Core.CellValue DrawingScale { get; set; }
        public Core.CellValue DrawingScaleType { get; set; }
        public Core.CellValue DrawingSizeType { get; set; }
        public Core.CellValue InhibitSnap { get; set; }
        public Core.CellValue Height { get; set; }
        public Core.CellValue Scale { get; set; }
        public Core.CellValue Width { get; set; }
        public Core.CellValue ShadowObliqueAngle { get; set; }
        public Core.CellValue ShadowOffsetX { get; set; }
        public Core.CellValue ShadowOffsetY { get; set; }
        public Core.CellValue ShadowScaleFactor { get; set; }
        public Core.CellValue ShadowType { get; set; }
        public Core.CellValue UIVisibility { get; set; }
        public Core.CellValue DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.DrawingScale), Core.SrcConstants.PageDrawingScale, this.DrawingScale);
            yield return this._create(nameof(this.DrawingScaleType), Core.SrcConstants.PageDrawingScaleType,
                this.DrawingScaleType);
            yield return this._create(nameof(this.DrawingSizeType), Core.SrcConstants.PageDrawingSizeType,
                this.DrawingSizeType);
            yield return this._create(nameof(this.InhibitSnap), Core.SrcConstants.PageInhibitSnap, this.InhibitSnap);
            yield return this._create(nameof(this.Height), Core.SrcConstants.PageHeight, this.Height);
            yield return this._create(nameof(this.Scale), Core.SrcConstants.PageScale, this.Scale);
            yield return this._create(nameof(this.Width), Core.SrcConstants.PageWidth, this.Width);
            yield return this._create(nameof(this.ShadowObliqueAngle), Core.SrcConstants.PageShadowObliqueAngle,
                this.ShadowObliqueAngle);
            yield return this._create(nameof(this.ShadowOffsetX), Core.SrcConstants.PageShadowOffsetX,
                this.ShadowOffsetX);
            yield return this._create(nameof(this.ShadowOffsetY), Core.SrcConstants.PageShadowOffsetY,
                this.ShadowOffsetY);
            yield return this._create(nameof(this.ShadowScaleFactor), Core.SrcConstants.PageShadowScaleFactor,
                this.ShadowScaleFactor);
            yield return this._create(nameof(this.ShadowType), Core.SrcConstants.PageShadowType, this.ShadowType);
            yield return this._create(nameof(this.UIVisibility), Core.SrcConstants.PageUIVisibility, this.UIVisibility);
            yield return this._create(nameof(this.DrawingResizeType), Core.SrcConstants.PageDrawingResizeType,
                this.DrawingResizeType);
        }


        public static PageFormatCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        public static PageFormatCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
        {
            var record = new PageFormatCells();
            var getcellvalue = getvalfromrowfunc(row, cols);

            record.DrawingScale = getcellvalue(nameof(DrawingScale));
            record.DrawingScaleType = getcellvalue(nameof(DrawingScaleType));
            record.DrawingSizeType = getcellvalue(nameof(DrawingSizeType));
            record.InhibitSnap = getcellvalue(nameof(InhibitSnap));
            record.Height = getcellvalue(nameof(Height));
            record.Scale = getcellvalue(nameof(Scale));
            record.Width = getcellvalue(nameof(Width));
            record.ShadowObliqueAngle = getcellvalue(nameof(ShadowObliqueAngle));
            record.ShadowOffsetX = getcellvalue(nameof(ShadowOffsetX));
            record.ShadowOffsetY = getcellvalue(nameof(ShadowOffsetY));
            record.ShadowScaleFactor = getcellvalue(nameof(ShadowScaleFactor));
            record.ShadowType = getcellvalue(nameof(ShadowType));
            record.UIVisibility = getcellvalue(nameof(UIVisibility));
            record.DrawingResizeType = getcellvalue(nameof(DrawingResizeType));

            return record;
        }

        class Builder : CellRecordBuilderCellQuery<PageFormatCells>
        {
            public Builder() : base(PageFormatCells.RowToRecord)
            {
            }

        }
    }
}