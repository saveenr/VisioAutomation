using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;

using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class FormatCells : VASS.CellGroups.CellGroup
    {
        public VisioAutomation.Core.CellValue DrawingScale { get; set; }
        public VisioAutomation.Core.CellValue DrawingScaleType { get; set; }
        public VisioAutomation.Core.CellValue DrawingSizeType { get; set; }
        public VisioAutomation.Core.CellValue InhibitSnap { get; set; }
        public VisioAutomation.Core.CellValue Height { get; set; }
        public VisioAutomation.Core.CellValue Scale { get; set; }
        public VisioAutomation.Core.CellValue Width { get; set; }
        public VisioAutomation.Core.CellValue ShadowObliqueAngle { get; set; }
        public VisioAutomation.Core.CellValue ShadowOffsetX { get; set; }
        public VisioAutomation.Core.CellValue ShadowOffsetY { get; set; }
        public VisioAutomation.Core.CellValue ShadowScaleFactor { get; set; }
        public VisioAutomation.Core.CellValue ShadowType { get; set; }
        public VisioAutomation.Core.CellValue UIVisibility { get; set; }
        public VisioAutomation.Core.CellValue DrawingResizeType { get; set; } // new in visio 2010

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.DrawingScale), VisioAutomation.Core.SrcConstants.PageDrawingScale, this.DrawingScale);
            yield return this.Create(nameof(this.DrawingScaleType), VisioAutomation.Core.SrcConstants.PageDrawingScaleType,
                this.DrawingScaleType);
            yield return this.Create(nameof(this.DrawingSizeType), VisioAutomation.Core.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
            yield return this.Create(nameof(this.InhibitSnap), VisioAutomation.Core.SrcConstants.PageInhibitSnap, this.InhibitSnap);
            yield return this.Create(nameof(this.Height), VisioAutomation.Core.SrcConstants.PageHeight, this.Height);
            yield return this.Create(nameof(this.Scale), VisioAutomation.Core.SrcConstants.PageScale, this.Scale);
            yield return this.Create(nameof(this.Width), VisioAutomation.Core.SrcConstants.PageWidth, this.Width);
            yield return this.Create(nameof(this.ShadowObliqueAngle), VisioAutomation.Core.SrcConstants.PageShadowObliqueAngle,
                this.ShadowObliqueAngle);
            yield return this.Create(nameof(this.ShadowOffsetX), VisioAutomation.Core.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
            yield return this.Create(nameof(this.ShadowOffsetY), VisioAutomation.Core.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
            yield return this.Create(nameof(this.ShadowScaleFactor), VisioAutomation.Core.SrcConstants.PageShadowScaleFactor,
                this.ShadowScaleFactor);
            yield return this.Create(nameof(this.ShadowType), VisioAutomation.Core.SrcConstants.PageShadowType, this.ShadowType);
            yield return this.Create(nameof(this.UIVisibility), VisioAutomation.Core.SrcConstants.PageUIVisibility, this.UIVisibility);
            yield return this.Create(nameof(this.DrawingResizeType), VisioAutomation.Core.SrcConstants.PageDrawingResizeType,
                this.DrawingResizeType);
        }


        public static FormatCells GetCells(IVisio.Shape shape, VisioAutomation.Core.CellValueType type)
        {
            var reader = PageFormatCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageFormatCellsBuilder> PageFormatCells_lazy_builder = new System.Lazy<PageFormatCellsBuilder>();

        class PageFormatCellsBuilder : VASS.CellGroups.CellGroupBuilder<FormatCells>
        {
            public PageFormatCellsBuilder() : base(VASS.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override FormatCells ToCellGroup(ShapeSheet.Query.Row<string> row, VisioAutomation.ShapeSheet.Query.Columns cols)
            {
                var cells = new FormatCells();
                var getcellvalue = VASS.CellGroups.CellGroup.row_to_cellgroup(row, cols);

                cells.DrawingScale = getcellvalue(nameof(FormatCells.DrawingScale));
                cells.DrawingScaleType = getcellvalue(nameof(FormatCells.DrawingScaleType));
                cells.DrawingSizeType = getcellvalue(nameof(FormatCells.DrawingSizeType));
                cells.InhibitSnap = getcellvalue(nameof(FormatCells.InhibitSnap));
                cells.Height = getcellvalue(nameof(FormatCells.Height));
                cells.Scale = getcellvalue(nameof(FormatCells.Scale));
                cells.Width = getcellvalue(nameof(FormatCells.Width));
                cells.ShadowObliqueAngle = getcellvalue(nameof(FormatCells.ShadowObliqueAngle));
                cells.ShadowOffsetX = getcellvalue(nameof(FormatCells.ShadowOffsetX));
                cells.ShadowOffsetY = getcellvalue(nameof(FormatCells.ShadowOffsetY));
                cells.ShadowScaleFactor = getcellvalue(nameof(FormatCells.ShadowScaleFactor));
                cells.ShadowType = getcellvalue(nameof(FormatCells.ShadowType));
                cells.UIVisibility = getcellvalue(nameof(FormatCells.UIVisibility));
                cells.DrawingResizeType = getcellvalue(nameof(FormatCells.DrawingResizeType));

                return cells;
            }
        }

    }
}