using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class FormatCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue FillBackground { get; set; }
        public VisioAutomation.Core.CellValue FillBackgroundTransparency { get; set; }
        public VisioAutomation.Core.CellValue FillForeground { get; set; }
        public VisioAutomation.Core.CellValue FillForegroundTransparency { get; set; }
        public VisioAutomation.Core.CellValue FillPattern { get; set; }
        public VisioAutomation.Core.CellValue FillShadowObliqueAngle { get; set; }
        public VisioAutomation.Core.CellValue FillShadowOffsetX { get; set; }
        public VisioAutomation.Core.CellValue FillShadowOffsetY { get; set; }
        public VisioAutomation.Core.CellValue FillShadowScaleFactor { get; set; }
        public VisioAutomation.Core.CellValue FillShadowType { get; set; }
        public VisioAutomation.Core.CellValue FillShadowBackground { get; set; }
        public VisioAutomation.Core.CellValue FillShadowBackgroundTransparency { get; set; }
        public VisioAutomation.Core.CellValue FillShadowForeground { get; set; }
        public VisioAutomation.Core.CellValue FillShadowForegroundTransparency { get; set; }
        public VisioAutomation.Core.CellValue FillShadowPattern { get; set; }
        public VisioAutomation.Core.CellValue LineBeginArrow { get; set; }
        public VisioAutomation.Core.CellValue LineBeginArrowSize { get; set; }
        public VisioAutomation.Core.CellValue LineEndArrow { get; set; }
        public VisioAutomation.Core.CellValue LineEndArrowSize { get; set; }
        public VisioAutomation.Core.CellValue LineCap { get; set; }
        public VisioAutomation.Core.CellValue LineColor { get; set; }
        public VisioAutomation.Core.CellValue LineColorTransparency { get; set; }
        public VisioAutomation.Core.CellValue LinePattern { get; set; }
        public VisioAutomation.Core.CellValue LineWeight { get; set; }
        public VisioAutomation.Core.CellValue LineRounding { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.FillBackground), VisioAutomation.Core.SrcConstants.FillBackground, this.FillBackground);
            yield return this.Create(nameof(this.FillBackgroundTransparency), VisioAutomation.Core.SrcConstants.FillBackgroundTransparency,
                this.FillBackgroundTransparency);
            yield return this.Create(nameof(this.FillForeground), VisioAutomation.Core.SrcConstants.FillForeground, this.FillForeground);
            yield return this.Create(nameof(this.FillForegroundTransparency), VisioAutomation.Core.SrcConstants.FillForegroundTransparency,
                this.FillForegroundTransparency);
            yield return this.Create(nameof(this.FillPattern), VisioAutomation.Core.SrcConstants.FillPattern, this.FillPattern);
            yield return this.Create(nameof(this.FillShadowObliqueAngle), VisioAutomation.Core.SrcConstants.FillShadowObliqueAngle,
                this.FillShadowObliqueAngle);
            yield return this.Create(nameof(this.FillShadowOffsetX), VisioAutomation.Core.SrcConstants.FillShadowOffsetX,
                this.FillShadowOffsetX);
            yield return this.Create(nameof(this.FillShadowOffsetY), VisioAutomation.Core.SrcConstants.FillShadowOffsetY,
                this.FillShadowOffsetY);
            yield return this.Create(nameof(this.FillShadowScaleFactor), VisioAutomation.Core.SrcConstants.FillShadowScaleFactor,
                this.FillShadowScaleFactor);
            yield return this.Create(nameof(this.FillShadowType), VisioAutomation.Core.SrcConstants.FillShadowType, this.FillShadowType);
            yield return this.Create(nameof(this.FillShadowBackground), VisioAutomation.Core.SrcConstants.FillShadowBackground,
                this.FillShadowBackground);
            yield return this.Create(nameof(this.FillShadowBackgroundTransparency),
                VisioAutomation.Core.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return this.Create(nameof(this.FillShadowForeground), VisioAutomation.Core.SrcConstants.FillShadowForeground,
                this.FillShadowForeground);
            yield return this.Create(nameof(this.FillShadowForegroundTransparency),
                VisioAutomation.Core.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return this.Create(nameof(this.FillShadowPattern), VisioAutomation.Core.SrcConstants.FillShadowPattern,
                this.FillShadowPattern);
            yield return this.Create(nameof(this.LineBeginArrow), VisioAutomation.Core.SrcConstants.LineBeginArrow, this.LineBeginArrow);
            yield return this.Create(nameof(this.LineBeginArrowSize), VisioAutomation.Core.SrcConstants.LineBeginArrowSize,
                this.LineBeginArrowSize);
            yield return this.Create(nameof(this.LineEndArrow), VisioAutomation.Core.SrcConstants.LineEndArrow, this.LineEndArrow);
            yield return this.Create(nameof(this.LineEndArrowSize), VisioAutomation.Core.SrcConstants.LineEndArrowSize, this.LineEndArrowSize);
            yield return this.Create(nameof(this.LineCap), VisioAutomation.Core.SrcConstants.LineCap, this.LineCap);
            yield return this.Create(nameof(this.LineColor), VisioAutomation.Core.SrcConstants.LineColor, this.LineColor);
            yield return this.Create(nameof(this.LineColorTransparency), VisioAutomation.Core.SrcConstants.LineColorTransparency,
                this.LineColorTransparency);
            yield return this.Create(nameof(this.LinePattern), VisioAutomation.Core.SrcConstants.LinePattern, this.LinePattern);
            yield return this.Create(nameof(this.LineWeight), VisioAutomation.Core.SrcConstants.LineWeight, this.LineWeight);
            yield return this.Create(nameof(this.LineRounding), VisioAutomation.Core.SrcConstants.LineRounding, this.LineRounding);
        }


        public static List<FormatCells> GetCells(IVisio.Page page, IList<int> shapeids, VisioAutomation.Core.CellValueType type)
        {
            var reader = shape_format_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static FormatCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = shape_format_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeFormatCellsBuilder> shape_format_lazy_builder = new System.Lazy<ShapeFormatCellsBuilder>();

        class ShapeFormatCellsBuilder : VASS.CellGroups.CellGroupBuilder<FormatCells>
        {

            public ShapeFormatCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override FormatCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {

                var cells = new FormatCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.FillBackground = getcellvalue(nameof(FormatCells.FillBackground));
                cells.FillBackgroundTransparency = getcellvalue(nameof(FormatCells.FillBackgroundTransparency));
                cells.FillForeground = getcellvalue(nameof(FormatCells.FillForeground));
                cells.FillForegroundTransparency = getcellvalue(nameof(FormatCells.FillForegroundTransparency));
                cells.FillPattern = getcellvalue(nameof(FormatCells.FillPattern));
                cells.FillShadowObliqueAngle = getcellvalue(nameof(FormatCells.FillShadowObliqueAngle));
                cells.FillShadowOffsetX = getcellvalue(nameof(FormatCells.FillShadowOffsetX));
                cells.FillShadowOffsetY = getcellvalue(nameof(FormatCells.FillShadowOffsetY));
                cells.FillShadowScaleFactor = getcellvalue(nameof(FormatCells.FillShadowScaleFactor));
                cells.FillShadowType = getcellvalue(nameof(FormatCells.FillShadowType));
                cells.FillShadowBackground = getcellvalue(nameof(FormatCells.FillShadowBackground));
                cells.FillShadowBackgroundTransparency = getcellvalue(nameof(FormatCells.FillShadowBackgroundTransparency));
                cells.FillShadowForeground = getcellvalue(nameof(FormatCells.FillShadowForeground));
                cells.FillShadowForegroundTransparency = getcellvalue(nameof(FormatCells.FillShadowForegroundTransparency));
                cells.FillShadowPattern = getcellvalue(nameof(FormatCells.FillShadowPattern));
                cells.LineBeginArrow = getcellvalue(nameof(FormatCells.LineBeginArrow));
                cells.LineBeginArrowSize = getcellvalue(nameof(FormatCells.LineBeginArrowSize));
                cells.LineEndArrow = getcellvalue(nameof(FormatCells.LineEndArrow));
                cells.LineEndArrowSize = getcellvalue(nameof(FormatCells.LineEndArrowSize));
                cells.LineCap = getcellvalue(nameof(FormatCells.LineCap));
                cells.LineColor = getcellvalue(nameof(FormatCells.LineColor));
                cells.LineColorTransparency = getcellvalue(nameof(FormatCells.LineColorTransparency));
                cells.LinePattern = getcellvalue(nameof(FormatCells.LinePattern));
                cells.LineWeight = getcellvalue(nameof(FormatCells.LineWeight));
                cells.LineRounding = getcellvalue(nameof(FormatCells.LineRounding));
                return cells;
            }

        }

    }
}

