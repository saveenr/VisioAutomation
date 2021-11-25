﻿using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Shapes
{
    public class ShapeFormatCells : VASS.CellGroups.CellGroup
    {
        public VASS.CellValue FillBackground { get; set; }
        public VASS.CellValue FillBackgroundTransparency { get; set; }
        public VASS.CellValue FillForeground { get; set; }
        public VASS.CellValue FillForegroundTransparency { get; set; }
        public VASS.CellValue FillPattern { get; set; }
        public VASS.CellValue FillShadowObliqueAngle { get; set; }
        public VASS.CellValue FillShadowOffsetX { get; set; }
        public VASS.CellValue FillShadowOffsetY { get; set; }
        public VASS.CellValue FillShadowScaleFactor { get; set; }
        public VASS.CellValue FillShadowType { get; set; }
        public VASS.CellValue FillShadowBackground { get; set; }
        public VASS.CellValue FillShadowBackgroundTransparency { get; set; }
        public VASS.CellValue FillShadowForeground { get; set; }
        public VASS.CellValue FillShadowForegroundTransparency { get; set; }
        public VASS.CellValue FillShadowPattern { get; set; }
        public VASS.CellValue LineBeginArrow { get; set; }
        public VASS.CellValue LineBeginArrowSize { get; set; }
        public VASS.CellValue LineEndArrow { get; set; }
        public VASS.CellValue LineEndArrowSize { get; set; }
        public VASS.CellValue LineCap { get; set; }
        public VASS.CellValue LineColor { get; set; }
        public VASS.CellValue LineColorTransparency { get; set; }
        public VASS.CellValue LinePattern { get; set; }
        public VASS.CellValue LineWeight { get; set; }
        public VASS.CellValue LineRounding { get; set; }

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.FillBackground), VASS.SrcConstants.FillBackground, this.FillBackground);
            yield return this.Create(nameof(this.FillBackgroundTransparency), VASS.SrcConstants.FillBackgroundTransparency,
                this.FillBackgroundTransparency);
            yield return this.Create(nameof(this.FillForeground), VASS.SrcConstants.FillForeground, this.FillForeground);
            yield return this.Create(nameof(this.FillForegroundTransparency), VASS.SrcConstants.FillForegroundTransparency,
                this.FillForegroundTransparency);
            yield return this.Create(nameof(this.FillPattern), VASS.SrcConstants.FillPattern, this.FillPattern);
            yield return this.Create(nameof(this.FillShadowObliqueAngle), VASS.SrcConstants.FillShadowObliqueAngle,
                this.FillShadowObliqueAngle);
            yield return this.Create(nameof(this.FillShadowOffsetX), VASS.SrcConstants.FillShadowOffsetX,
                this.FillShadowOffsetX);
            yield return this.Create(nameof(this.FillShadowOffsetY), VASS.SrcConstants.FillShadowOffsetY,
                this.FillShadowOffsetY);
            yield return this.Create(nameof(this.FillShadowScaleFactor), VASS.SrcConstants.FillShadowScaleFactor,
                this.FillShadowScaleFactor);
            yield return this.Create(nameof(this.FillShadowType), VASS.SrcConstants.FillShadowType, this.FillShadowType);
            yield return this.Create(nameof(this.FillShadowBackground), VASS.SrcConstants.FillShadowBackground,
                this.FillShadowBackground);
            yield return this.Create(nameof(this.FillShadowBackgroundTransparency),
                VASS.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
            yield return this.Create(nameof(this.FillShadowForeground), VASS.SrcConstants.FillShadowForeground,
                this.FillShadowForeground);
            yield return this.Create(nameof(this.FillShadowForegroundTransparency),
                VASS.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
            yield return this.Create(nameof(this.FillShadowPattern), VASS.SrcConstants.FillShadowPattern,
                this.FillShadowPattern);
            yield return this.Create(nameof(this.LineBeginArrow), VASS.SrcConstants.LineBeginArrow, this.LineBeginArrow);
            yield return this.Create(nameof(this.LineBeginArrowSize), VASS.SrcConstants.LineBeginArrowSize,
                this.LineBeginArrowSize);
            yield return this.Create(nameof(this.LineEndArrow), VASS.SrcConstants.LineEndArrow, this.LineEndArrow);
            yield return this.Create(nameof(this.LineEndArrowSize), VASS.SrcConstants.LineEndArrowSize, this.LineEndArrowSize);
            yield return this.Create(nameof(this.LineCap), VASS.SrcConstants.LineCap, this.LineCap);
            yield return this.Create(nameof(this.LineColor), VASS.SrcConstants.LineColor, this.LineColor);
            yield return this.Create(nameof(this.LineColorTransparency), VASS.SrcConstants.LineColorTransparency,
                this.LineColorTransparency);
            yield return this.Create(nameof(this.LinePattern), VASS.SrcConstants.LinePattern, this.LinePattern);
            yield return this.Create(nameof(this.LineWeight), VASS.SrcConstants.LineWeight, this.LineWeight);
            yield return this.Create(nameof(this.LineRounding), VASS.SrcConstants.LineRounding, this.LineRounding);
        }


        public static List<ShapeFormatCells> GetCells(IVisio.Page page, IList<int> shapeids, VASS.CellValueType type)
        {
            var reader = shape_format_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeFormatCells GetCells(IVisio.Shape shape, VASS.CellValueType type)
        {
            var reader = shape_format_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeFormatCellsBuilder> shape_format_lazy_builder = new System.Lazy<ShapeFormatCellsBuilder>();

        class ShapeFormatCellsBuilder : VASS.CellGroups.CellGroupBuilder<ShapeFormatCells>
        {

            public ShapeFormatCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override ShapeFormatCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {

                var cells = new ShapeFormatCells();
                var getcellvalue = VisioAutomation.ShapeSheet.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.FillBackground = getcellvalue(nameof(ShapeFormatCells.FillBackground));
                cells.FillBackgroundTransparency = getcellvalue(nameof(ShapeFormatCells.FillBackgroundTransparency));
                cells.FillForeground = getcellvalue(nameof(ShapeFormatCells.FillForeground));
                cells.FillForegroundTransparency = getcellvalue(nameof(ShapeFormatCells.FillForegroundTransparency));
                cells.FillPattern = getcellvalue(nameof(ShapeFormatCells.FillPattern));
                cells.FillShadowObliqueAngle = getcellvalue(nameof(ShapeFormatCells.FillShadowObliqueAngle));
                cells.FillShadowOffsetX = getcellvalue(nameof(ShapeFormatCells.FillShadowOffsetX));
                cells.FillShadowOffsetY = getcellvalue(nameof(ShapeFormatCells.FillShadowOffsetY));
                cells.FillShadowScaleFactor = getcellvalue(nameof(ShapeFormatCells.FillShadowScaleFactor));
                cells.FillShadowType = getcellvalue(nameof(ShapeFormatCells.FillShadowType));
                cells.FillShadowBackground = getcellvalue(nameof(ShapeFormatCells.FillShadowBackground));
                cells.FillShadowBackgroundTransparency = getcellvalue(nameof(ShapeFormatCells.FillShadowBackgroundTransparency));
                cells.FillShadowForeground = getcellvalue(nameof(ShapeFormatCells.FillShadowForeground));
                cells.FillShadowForegroundTransparency = getcellvalue(nameof(ShapeFormatCells.FillShadowForegroundTransparency));
                cells.FillShadowPattern = getcellvalue(nameof(ShapeFormatCells.FillShadowPattern));
                cells.LineBeginArrow = getcellvalue(nameof(ShapeFormatCells.LineBeginArrow));
                cells.LineBeginArrowSize = getcellvalue(nameof(ShapeFormatCells.LineBeginArrowSize));
                cells.LineEndArrow = getcellvalue(nameof(ShapeFormatCells.LineEndArrow));
                cells.LineEndArrowSize = getcellvalue(nameof(ShapeFormatCells.LineEndArrowSize));
                cells.LineCap = getcellvalue(nameof(ShapeFormatCells.LineCap));
                cells.LineColor = getcellvalue(nameof(ShapeFormatCells.LineColor));
                cells.LineColorTransparency = getcellvalue(nameof(ShapeFormatCells.LineColorTransparency));
                cells.LinePattern = getcellvalue(nameof(ShapeFormatCells.LinePattern));
                cells.LineWeight = getcellvalue(nameof(ShapeFormatCells.LineWeight));
                cells.LineRounding = getcellvalue(nameof(ShapeFormatCells.LineRounding));
                return cells;
            }

        }

    }
}

