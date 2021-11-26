using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class FormatCells : CellGroup
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

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.FillBackground), Core.SrcConstants.FillBackground, this.FillBackground);
            yield return this.Create(nameof(this.FillBackgroundTransparency), Core.SrcConstants.FillBackgroundTransparency,
                this.FillBackgroundTransparency);
            yield return this.Create(nameof(this.FillForeground), Core.SrcConstants.FillForeground, this.FillForeground);
            yield return this.Create(nameof(this.FillForegroundTransparency), Core.SrcConstants.FillForegroundTransparency,
                this.FillForegroundTransparency);
            yield return this.Create(nameof(this.FillPattern), Core.SrcConstants.FillPattern, this.FillPattern);
            yield return this.Create(nameof(this.FillShadowObliqueAngle), Core.SrcConstants.FillShadowObliqueAngle,
                this.FillShadowObliqueAngle);
            yield return this.Create(nameof(this.FillShadowOffsetX), Core.SrcConstants.FillShadowOffsetX,
                this.FillShadowOffsetX);
            yield return this.Create(nameof(this.FillShadowOffsetY), Core.SrcConstants.FillShadowOffsetY,
                this.FillShadowOffsetY);
            yield return this.Create(nameof(this.FillShadowScaleFactor), Core.SrcConstants.FillShadowScaleFactor,
                this.FillShadowScaleFactor);
            yield return this.Create(nameof(this.FillShadowType), Core.SrcConstants.FillShadowType, this.FillShadowType);
            yield return this.Create(nameof(this.FillShadowBackground), Core.SrcConstants.FillShadowBackground,
                this.FillShadowBackground);
            yield return this.Create(nameof(this.FillShadowBackgroundTransparency),
                Core.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return this.Create(nameof(this.FillShadowForeground), Core.SrcConstants.FillShadowForeground,
                this.FillShadowForeground);
            yield return this.Create(nameof(this.FillShadowForegroundTransparency),
                Core.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return this.Create(nameof(this.FillShadowPattern), Core.SrcConstants.FillShadowPattern,
                this.FillShadowPattern);
            yield return this.Create(nameof(this.LineBeginArrow), Core.SrcConstants.LineBeginArrow, this.LineBeginArrow);
            yield return this.Create(nameof(this.LineBeginArrowSize), Core.SrcConstants.LineBeginArrowSize,
                this.LineBeginArrowSize);
            yield return this.Create(nameof(this.LineEndArrow), Core.SrcConstants.LineEndArrow, this.LineEndArrow);
            yield return this.Create(nameof(this.LineEndArrowSize), Core.SrcConstants.LineEndArrowSize, this.LineEndArrowSize);
            yield return this.Create(nameof(this.LineCap), Core.SrcConstants.LineCap, this.LineCap);
            yield return this.Create(nameof(this.LineColor), Core.SrcConstants.LineColor, this.LineColor);
            yield return this.Create(nameof(this.LineColorTransparency), Core.SrcConstants.LineColorTransparency,
                this.LineColorTransparency);
            yield return this.Create(nameof(this.LinePattern), Core.SrcConstants.LinePattern, this.LinePattern);
            yield return this.Create(nameof(this.LineWeight), Core.SrcConstants.LineWeight, this.LineWeight);
            yield return this.Create(nameof(this.LineRounding), Core.SrcConstants.LineRounding, this.LineRounding);
        }


        public static List<FormatCells> GetCells(IVisio.Page page, IList<int> shapeids, Core.CellValueType type)
        {
            var reader = shape_format_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static FormatCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = shape_format_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeFormatCellsBuilder> shape_format_lazy_builder = new System.Lazy<ShapeFormatCellsBuilder>();

        class ShapeFormatCellsBuilder : CellGroupBuilder<FormatCells>
        {

            public ShapeFormatCellsBuilder() : base(CellGroupBuilderType.SingleRow)
            {
            }

            public override FormatCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {

                var cells = new FormatCells();
                var getcellvalue = row_to_cellgroup(row, cols);

                cells.FillBackground = getcellvalue(nameof(FillBackground));
                cells.FillBackgroundTransparency = getcellvalue(nameof(FillBackgroundTransparency));
                cells.FillForeground = getcellvalue(nameof(FillForeground));
                cells.FillForegroundTransparency = getcellvalue(nameof(FillForegroundTransparency));
                cells.FillPattern = getcellvalue(nameof(FillPattern));
                cells.FillShadowObliqueAngle = getcellvalue(nameof(FillShadowObliqueAngle));
                cells.FillShadowOffsetX = getcellvalue(nameof(FillShadowOffsetX));
                cells.FillShadowOffsetY = getcellvalue(nameof(FillShadowOffsetY));
                cells.FillShadowScaleFactor = getcellvalue(nameof(FillShadowScaleFactor));
                cells.FillShadowType = getcellvalue(nameof(FillShadowType));
                cells.FillShadowBackground = getcellvalue(nameof(FillShadowBackground));
                cells.FillShadowBackgroundTransparency = getcellvalue(nameof(FillShadowBackgroundTransparency));
                cells.FillShadowForeground = getcellvalue(nameof(FillShadowForeground));
                cells.FillShadowForegroundTransparency = getcellvalue(nameof(FillShadowForegroundTransparency));
                cells.FillShadowPattern = getcellvalue(nameof(FillShadowPattern));
                cells.LineBeginArrow = getcellvalue(nameof(LineBeginArrow));
                cells.LineBeginArrowSize = getcellvalue(nameof(LineBeginArrowSize));
                cells.LineEndArrow = getcellvalue(nameof(LineEndArrow));
                cells.LineEndArrowSize = getcellvalue(nameof(LineEndArrowSize));
                cells.LineCap = getcellvalue(nameof(LineCap));
                cells.LineColor = getcellvalue(nameof(LineColor));
                cells.LineColorTransparency = getcellvalue(nameof(LineColorTransparency));
                cells.LinePattern = getcellvalue(nameof(LinePattern));
                cells.LineWeight = getcellvalue(nameof(LineWeight));
                cells.LineRounding = getcellvalue(nameof(LineRounding));
                return cells;
            }

        }

    }
}

