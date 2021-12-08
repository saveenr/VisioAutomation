using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellRecords;
using VASS = VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeFormatCells : CellRecord
    {
        public Core.CellValue FillBackground { get; set; }
        public Core.CellValue FillBackgroundTransparency { get; set; }
        public Core.CellValue FillForeground { get; set; }
        public Core.CellValue FillForegroundTransparency { get; set; }
        public Core.CellValue FillPattern { get; set; }
        public Core.CellValue FillShadowObliqueAngle { get; set; }
        public Core.CellValue FillShadowOffsetX { get; set; }
        public Core.CellValue FillShadowOffsetY { get; set; }
        public Core.CellValue FillShadowScaleFactor { get; set; }
        public Core.CellValue FillShadowType { get; set; }
        public Core.CellValue FillShadowBackground { get; set; }
        public Core.CellValue FillShadowBackgroundTransparency { get; set; }
        public Core.CellValue FillShadowForeground { get; set; }
        public Core.CellValue FillShadowForegroundTransparency { get; set; }
        public Core.CellValue FillShadowPattern { get; set; }
        public Core.CellValue LineBeginArrow { get; set; }
        public Core.CellValue LineBeginArrowSize { get; set; }
        public Core.CellValue LineEndArrow { get; set; }
        public Core.CellValue LineEndArrowSize { get; set; }
        public Core.CellValue LineCap { get; set; }
        public Core.CellValue LineColor { get; set; }
        public Core.CellValue LineColorTransparency { get; set; }
        public Core.CellValue LinePattern { get; set; }
        public Core.CellValue LineWeight { get; set; }
        public Core.CellValue LineRounding { get; set; }

        public override IEnumerable<CellMetadata> GetCellMetadata()
        {
            yield return this._create(nameof(this.FillBackground), Core.SrcConstants.FillBackground,
                this.FillBackground);
            yield return this._create(nameof(this.FillBackgroundTransparency),
                Core.SrcConstants.FillBackgroundTransparency,
                this.FillBackgroundTransparency);
            yield return this._create(nameof(this.FillForeground), Core.SrcConstants.FillForeground,
                this.FillForeground);
            yield return this._create(nameof(this.FillForegroundTransparency),
                Core.SrcConstants.FillForegroundTransparency,
                this.FillForegroundTransparency);
            yield return this._create(nameof(this.FillPattern), Core.SrcConstants.FillPattern, this.FillPattern);
            yield return this._create(nameof(this.FillShadowObliqueAngle), Core.SrcConstants.FillShadowObliqueAngle,
                this.FillShadowObliqueAngle);
            yield return this._create(nameof(this.FillShadowOffsetX), Core.SrcConstants.FillShadowOffsetX,
                this.FillShadowOffsetX);
            yield return this._create(nameof(this.FillShadowOffsetY), Core.SrcConstants.FillShadowOffsetY,
                this.FillShadowOffsetY);
            yield return this._create(nameof(this.FillShadowScaleFactor), Core.SrcConstants.FillShadowScaleFactor,
                this.FillShadowScaleFactor);
            yield return this._create(nameof(this.FillShadowType), Core.SrcConstants.FillShadowType,
                this.FillShadowType);
            yield return this._create(nameof(this.FillShadowBackground), Core.SrcConstants.FillShadowBackground,
                this.FillShadowBackground);
            yield return this._create(nameof(this.FillShadowBackgroundTransparency),
                Core.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return this._create(nameof(this.FillShadowForeground), Core.SrcConstants.FillShadowForeground,
                this.FillShadowForeground);
            yield return this._create(nameof(this.FillShadowForegroundTransparency),
                Core.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return this._create(nameof(this.FillShadowPattern), Core.SrcConstants.FillShadowPattern,
                this.FillShadowPattern);
            yield return this._create(nameof(this.LineBeginArrow), Core.SrcConstants.LineBeginArrow,
                this.LineBeginArrow);
            yield return this._create(nameof(this.LineBeginArrowSize), Core.SrcConstants.LineBeginArrowSize,
                this.LineBeginArrowSize);
            yield return this._create(nameof(this.LineEndArrow), Core.SrcConstants.LineEndArrow, this.LineEndArrow);
            yield return this._create(nameof(this.LineEndArrowSize), Core.SrcConstants.LineEndArrowSize,
                this.LineEndArrowSize);
            yield return this._create(nameof(this.LineCap), Core.SrcConstants.LineCap, this.LineCap);
            yield return this._create(nameof(this.LineColor), Core.SrcConstants.LineColor, this.LineColor);
            yield return this._create(nameof(this.LineColorTransparency), Core.SrcConstants.LineColorTransparency,
                this.LineColorTransparency);
            yield return this._create(nameof(this.LinePattern), Core.SrcConstants.LinePattern, this.LinePattern);
            yield return this._create(nameof(this.LineWeight), Core.SrcConstants.LineWeight, this.LineWeight);
            yield return this._create(nameof(this.LineRounding), Core.SrcConstants.LineRounding, this.LineRounding);
        }


        public static List<ShapeFormatCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsMultipleShapesSingleRow(page, shapeids, type);
        }

        public static ShapeFormatCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = builder.Value;
            return reader.GetCellsSingleShapeSingleRow(shape, type);
        }

        private static readonly System.Lazy<Builder> builder = new System.Lazy<Builder>();

        class Builder : CellRecordBuilder<ShapeFormatCells>
        {
            public Builder() : base(CellRecordQueryType.CellQuery)
            {
            }

            public override ShapeFormatCells RowToRecord(VASS.Data.DataRow<string> row, VASS.Data.DataColumns cols)
            {
                var record = new ShapeFormatCells();
                var getcellvalue = queryrow_to_cellrecord(row, cols);

                record.FillBackground = getcellvalue(nameof(FillBackground));
                record.FillBackgroundTransparency = getcellvalue(nameof(FillBackgroundTransparency));
                record.FillForeground = getcellvalue(nameof(FillForeground));
                record.FillForegroundTransparency = getcellvalue(nameof(FillForegroundTransparency));
                record.FillPattern = getcellvalue(nameof(FillPattern));
                record.FillShadowObliqueAngle = getcellvalue(nameof(FillShadowObliqueAngle));
                record.FillShadowOffsetX = getcellvalue(nameof(FillShadowOffsetX));
                record.FillShadowOffsetY = getcellvalue(nameof(FillShadowOffsetY));
                record.FillShadowScaleFactor = getcellvalue(nameof(FillShadowScaleFactor));
                record.FillShadowType = getcellvalue(nameof(FillShadowType));
                record.FillShadowBackground = getcellvalue(nameof(FillShadowBackground));
                record.FillShadowBackgroundTransparency = getcellvalue(nameof(FillShadowBackgroundTransparency));
                record.FillShadowForeground = getcellvalue(nameof(FillShadowForeground));
                record.FillShadowForegroundTransparency = getcellvalue(nameof(FillShadowForegroundTransparency));
                record.FillShadowPattern = getcellvalue(nameof(FillShadowPattern));
                record.LineBeginArrow = getcellvalue(nameof(LineBeginArrow));
                record.LineBeginArrowSize = getcellvalue(nameof(LineBeginArrowSize));
                record.LineEndArrow = getcellvalue(nameof(LineEndArrow));
                record.LineEndArrowSize = getcellvalue(nameof(LineEndArrowSize));
                record.LineCap = getcellvalue(nameof(LineCap));
                record.LineColor = getcellvalue(nameof(LineColor));
                record.LineColorTransparency = getcellvalue(nameof(LineColorTransparency));
                record.LinePattern = getcellvalue(nameof(LinePattern));
                record.LineWeight = getcellvalue(nameof(LineWeight));
                record.LineRounding = getcellvalue(nameof(LineRounding));
                return record;
            }
        }
    }
}