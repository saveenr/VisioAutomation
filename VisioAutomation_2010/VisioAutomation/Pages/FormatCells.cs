using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VASS=VisioAutomation.ShapeSheet;

using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Pages
{
    public class FormatCells : CellGroup
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

        public override IEnumerable<CellMetadataItem> GetCellMetadata()
        {
            yield return this.Create(nameof(this.DrawingScale), Core.SrcConstants.PageDrawingScale, this.DrawingScale);
            yield return this.Create(nameof(this.DrawingScaleType), Core.SrcConstants.PageDrawingScaleType,
                this.DrawingScaleType);
            yield return this.Create(nameof(this.DrawingSizeType), Core.SrcConstants.PageDrawingSizeType, this.DrawingSizeType);
            yield return this.Create(nameof(this.InhibitSnap), Core.SrcConstants.PageInhibitSnap, this.InhibitSnap);
            yield return this.Create(nameof(this.Height), Core.SrcConstants.PageHeight, this.Height);
            yield return this.Create(nameof(this.Scale), Core.SrcConstants.PageScale, this.Scale);
            yield return this.Create(nameof(this.Width), Core.SrcConstants.PageWidth, this.Width);
            yield return this.Create(nameof(this.ShadowObliqueAngle), Core.SrcConstants.PageShadowObliqueAngle,
                this.ShadowObliqueAngle);
            yield return this.Create(nameof(this.ShadowOffsetX), Core.SrcConstants.PageShadowOffsetX, this.ShadowOffsetX);
            yield return this.Create(nameof(this.ShadowOffsetY), Core.SrcConstants.PageShadowOffsetY, this.ShadowOffsetY);
            yield return this.Create(nameof(this.ShadowScaleFactor), Core.SrcConstants.PageShadowScaleFactor,
                this.ShadowScaleFactor);
            yield return this.Create(nameof(this.ShadowType), Core.SrcConstants.PageShadowType, this.ShadowType);
            yield return this.Create(nameof(this.UIVisibility), Core.SrcConstants.PageUIVisibility, this.UIVisibility);
            yield return this.Create(nameof(this.DrawingResizeType), Core.SrcConstants.PageDrawingResizeType,
                this.DrawingResizeType);
        }


        public static FormatCells GetCells(IVisio.Shape shape, Core.CellValueType type)
        {
            var reader = PageFormatCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<PageFormatCellsBuilder> PageFormatCells_lazy_builder = new System.Lazy<PageFormatCellsBuilder>();

        class PageFormatCellsBuilder : CellGroupBuilder<FormatCells>
        {
            public PageFormatCellsBuilder() : base(CellGroupBuilderType.SingleRow)
            {
            }

            public override FormatCells ToCellGroup(VASS.Query.Row<string> row, VASS.Query.Columns cols)
            {
                var cells = new FormatCells();
                var getcellvalue = row_to_cellgroup(row, cols);

                cells.DrawingScale = getcellvalue(nameof(DrawingScale));
                cells.DrawingScaleType = getcellvalue(nameof(DrawingScaleType));
                cells.DrawingSizeType = getcellvalue(nameof(DrawingSizeType));
                cells.InhibitSnap = getcellvalue(nameof(InhibitSnap));
                cells.Height = getcellvalue(nameof(Height));
                cells.Scale = getcellvalue(nameof(Scale));
                cells.Width = getcellvalue(nameof(Width));
                cells.ShadowObliqueAngle = getcellvalue(nameof(ShadowObliqueAngle));
                cells.ShadowOffsetX = getcellvalue(nameof(ShadowOffsetX));
                cells.ShadowOffsetY = getcellvalue(nameof(ShadowOffsetY));
                cells.ShadowScaleFactor = getcellvalue(nameof(ShadowScaleFactor));
                cells.ShadowType = getcellvalue(nameof(ShadowType));
                cells.UIVisibility = getcellvalue(nameof(UIVisibility));
                cells.DrawingResizeType = getcellvalue(nameof(DrawingResizeType));

                return cells;
            }
        }

    }
}